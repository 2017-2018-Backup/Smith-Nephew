Option Strict Off
Option Explicit On
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports VB = Microsoft.VisualBasic

Public Class MIRBODYPROF
    Dim idLastCreated As ObjectId
    Function PR_GetBodyProfile() As Boolean
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim ptEntOpts As PromptEntityOptions = New PromptEntityOptions("Select Body Profile Polyline to Mirror:")
        ptEntOpts.AllowNone = True
        ptEntOpts.Keywords.Add("Spline")
        Dim ptEntRes As PromptEntityResult = ed.GetEntity(ptEntOpts)
        If ptEntRes.Status <> PromptStatus.OK Then
            MsgBox("Profile not selected", 16, "Mirror Body Profile")
            Return False
        End If
        Dim strProfile As String = ""
        Dim xyStart, xyEnd As ARMDIA1.XY
        Try
            Dim idSpline As ObjectId = ptEntRes.ObjectId
            Using tr As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acEnt As Spline = tr.GetObject(idSpline, OpenMode.ForRead)
                Dim rb As ResultBuffer = acEnt.GetXDataForApplication("ProfileID")
                If Not IsNothing(rb) Then
                    ' Get the values in the xdata
                    For Each typeVal As TypedValue In rb
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        strProfile = strProfile & typeVal.Value
                    Next
                    ''// Get the CenterLine
                    ''----------UserSelection("clear")
                    ''----------UserSelection("update")
                    ''--hSelectionCL = Open("selection", "DB ID = '" + sProfileID + "' AND  DB curvetype = 'CentreLine'")
                    Dim filterCenterLine(1) As TypedValue
                    filterCenterLine.SetValue(New TypedValue(DxfCode.Start, "Line"), 0)
                    filterCenterLine.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "curvetype"), 1)
                    ''filterCenterLine.SetValue(New TypedValue(DxfCode.ExtendedDataAsciiString, "CentreLine"), 2)
                    Dim selFilterCenLine As SelectionFilter = New SelectionFilter(filterCenterLine)
                    Dim selResultCenLine As PromptSelectionResult = ed.SelectAll(selFilterCenLine)
                    ''-------If (hSelectionCL) Then
                    If (selResultCenLine.Status = PromptStatus.OK) Then
                        Dim selSetCL As SelectionSet = selResultCenLine.Value
                        For Each idLine As ObjectId In selSetCL.GetObjectIds()
                            Dim acLine As Line = tr.GetObject(idLine, OpenMode.ForRead)
                            Dim ptStart As Point3d = acLine.StartPoint
                            Dim ptEnd As Point3d = acLine.EndPoint
                            ARMDIA1.PR_MakeXY(xyStart, ptStart.X, ptStart.Y)
                            ARMDIA1.PR_MakeXY(xyEnd, ptEnd.X, ptEnd.Y)
                        Next
                    Else
                        ''-----------Close ("selection", hSelectionCL )
                        MsgBox("Centre line NOT FOUND", 16, "Mirror Body Profile")
                        Return False
                    End If
                Else
                    MsgBox("Polyline picked NOT a Body Profile!", 16, "Mirror Body Profile")
                    Return False
                End If
                tr.Commit()
            End Using
        Catch ex As Exception
            MsgBox("Polyline picked NOT a Body Profile!", 16, "Mirror Body Profile")
            Return False
        End Try
        Dim filterType(5) As TypedValue
        filterType.SetValue(New TypedValue(DxfCode.Start, "Spline"), 0)
        filterType.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "curvetype"), 1)
        filterType.SetValue(New TypedValue(DxfCode.Operator, "<or"), 2)
        filterType.SetValue(New TypedValue(DxfCode.ExtendedDataAsciiString, "BackCurve"), 3)
        filterType.SetValue(New TypedValue(DxfCode.ExtendedDataAsciiString, "BackCurveLargest"), 4)
        filterType.SetValue(New TypedValue(DxfCode.Operator, "or>"), 5)
        Dim selFilter As SelectionFilter = New SelectionFilter(filterType)
        Dim selResult As PromptSelectionResult = ed.SelectAll(selFilter)
        If selResult.Status <> PromptStatus.OK Then
            Return False
        End If
        Dim selectionSet As SelectionSet = selResult.Value
        Dim ptCurveColl As Point3dCollection
        Dim sCurveLayer As String = ""
        Dim sSide As String = ""
        Dim nn As Integer = 0
        Dim fitPtsCount As Integer = 0
        For Each idObject As ObjectId In selectionSet.GetObjectIds()
            Using acTr As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acCurve As Spline = acTr.GetObject(idObject, OpenMode.ForRead)
                Dim resbuf As ResultBuffer = acCurve.GetXDataForApplication("Leg")
                If Not IsNothing(resbuf) Then
                    For Each typeVal As TypedValue In resbuf
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        sSide = sSide & typeVal.Value
                    Next
                End If
                fitPtsCount = acCurve.NumFitPoints
                nn = nn + 1
                If fitPtsCount < 2 Then
                    ''------Close("selection", hSelectionCurves)
                    MsgBox("Can't copy selected profile!" & Chr(10), 16, "Mirror Body Profile")
                    Return False
                End If
                ptCurveColl = New Point3dCollection
                Dim i As Integer = 0
                While (i < fitPtsCount)
                    Dim xyPt As ARMDIA1.XY
                    ARMDIA1.PR_MakeXY(xyPt, acCurve.GetFitPointAt(i).X, acCurve.GetFitPointAt(i).Y)
                    xyPt.y = xyStart.y - (xyPt.y - xyStart.y)
                    ptCurveColl.Add(New Point3d(xyPt.X, xyPt.y, 0))
                    i = i + 1
                End While
                sCurveLayer = acCurve.Layer
            End Using
            ARMDIA1.PR_SetLayer(sCurveLayer)
            PR_DrawSpline(ptCurveColl, strProfile)
            'If sSide.Contains("Left&Right") Then
            '    ARMDIA1.PR_SetLayer("Notes")
            '    Dim xyPt1 As ARMDIA1.XY
            '    ARMDIA1.PR_MakeXY(xyPt1, ptCurveColl(fitPtsCount - nn).X, ptCurveColl(fitPtsCount - nn).Y)
            '    PR_DrawClosedArrow(xyPt1, 90)
            '    xyPt1.y = xyPt1.y + 0.25
            '    ARMDIA1.PR_DrawText(sSide, xyPt1, 0.125, 0)
            'End If
        Next
        ''-----------Close("selection", hSelectionCurves)
        ''-----------Close("selection", hSelectionCL)
        Return True
    End Function
    'To Draw Spline
    Sub PR_DrawSpline(ByRef PointCollection As Point3dCollection, ByVal strProfileData As String)
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        '' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                                        OpenMode.ForWrite)
            '' Get a 3D vector from the point (0.5,0.5,0)
            Dim vecTan As Vector3d = New Point3d(0, 0, 0).GetAsVector
            '' Create a spline through (0, 0, 0), (5, 5, 0), and (10, 0, 0) with a
            '' start and end tangency of (0.5, 0.5, 0.0)
            Using acSpline As Spline = New Spline(PointCollection, vecTan, vecTan, 4, 0.0)
                'If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                '    acSpline.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyOtemplate.X, xyOtemplate.y, 0)))
                'End If
                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acSpline)
                acTrans.AddNewlyCreatedDBObject(acSpline, True)

                Dim acRegAppTbl As RegAppTable
                acRegAppTbl = acTrans.GetObject(acCurDb.RegAppTableId, OpenMode.ForRead)
                Dim acRegAppTblRec As RegAppTableRecord
                Dim appName As String = "CurveType"
                If acRegAppTbl.Has(appName) = False Then
                    acRegAppTblRec = New RegAppTableRecord
                    acRegAppTblRec.Name = appName
                    acRegAppTbl.UpgradeOpen()
                    acRegAppTbl.Add(acRegAppTblRec)
                    acTrans.AddNewlyCreatedDBObject(acRegAppTblRec, True)
                End If
                Using rb As New ResultBuffer
                    rb.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, appName))
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, strProfileData))
                    acSpline.XData = rb
                End Using
            End Using
            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub

    ''-----------------LGCNHEEL.D-----------------------------
    Public Sub PR_DrawHeelContracture()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim ptEntOpts As PromptEntityOptions = New PromptEntityOptions(Chr(10) & "Select a Leg Profile")
        ptEntOpts.AllowNone = True
        ''ptEntOpts.Keywords.Add("Spline")
        ptEntOpts.SetRejectMessage(Chr(10) & "Only Curve")
        ptEntOpts.AddAllowedClass(GetType(Spline), True)
        Dim ptEntRes As PromptEntityResult = ed.GetEntity(ptEntOpts)
        If ptEntRes.Status <> PromptStatus.OK Then
            MsgBox("Profile not selected", 16, "Heel Contracture")
            Exit Sub
        End If
        Dim strProfile As String = ""
        Dim ptLegInsertion As New Point3d
        Try
            Dim idSpline As ObjectId = ptEntRes.ObjectId
            Using tr As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acEnt As Spline = tr.GetObject(idSpline, OpenMode.ForRead)
                Dim rb As ResultBuffer = acEnt.GetXDataForApplication("ID")
                If Not IsNothing(rb) Then
                    ' Get the values in the xdata
                    For Each typeVal As TypedValue In rb
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        strProfile = strProfile & typeVal.Value
                    Next
                End If
                Dim rbInsertion As ResultBuffer = acEnt.GetXDataForApplication("Insertion")
                Dim strInsertion As String = ""
                If Not IsNothing(rbInsertion) Then
                    For Each typeVal As TypedValue In rbInsertion
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        strInsertion = strInsertion & typeVal.Value
                    Next
                End If
                If strInsertion <> "" Then
                    Dim xVal As Double = FN_GetNumber(strInsertion, 1)
                    Dim yVal As Double = FN_GetNumber(strInsertion, 2)
                    ptLegInsertion = New Point3d(xVal, yVal, 0)
                End If
            End Using
        Catch ex As Exception

        End Try
        strProfile = Trim(strProfile)
        Dim sStyleID, sLeg As String
        If strProfile.Contains("LeftLegCurve") Then
            sStyleID = Mid(strProfile, 1, strProfile.Length - 12)
            sLeg = "Left"
        ElseIf strProfile.Contains("RightLegCurve") Then
            sStyleID = Mid(strProfile, 1, strProfile.Length - 13)
            sLeg = "Right"
        Else
            MsgBox("A Leg Profile was not selected", 16, "Heel Contracture")
            Exit Sub
        End If

        Dim sOriginXdata As String = sStyleID + sLeg + "Origin"
        Dim filterType(1) As TypedValue
        filterType.SetValue(New TypedValue(DxfCode.BlockName, "X Marker"), 0)
        filterType.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "OriginID"), 1)
        Dim selFilter As SelectionFilter = New SelectionFilter(filterType)
        Dim selResult As PromptSelectionResult = ed.SelectAll(selFilter)
        If selResult.Status <> PromptStatus.OK Then
            MsgBox("A Leg Profile was not selected", 16, "Heel Contracture")
            Exit Sub
        End If
        Dim selectionSet As SelectionSet = selResult.Value
        Dim sHeelXdata As String = sStyleID + sLeg + "Heel"
        Dim HeelfilterType(1) As TypedValue
        HeelfilterType.SetValue(New TypedValue(DxfCode.BlockName, "X Marker"), 0)
        HeelfilterType.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "HeelID"), 1)
        Dim selHeelFilter As SelectionFilter = New SelectionFilter(HeelfilterType)
        selResult = ed.SelectAll(selHeelFilter)
        If selResult.Status <> PromptStatus.OK Then
            MsgBox("A Leg Profile was not selected", 16, "Heel Contracture")
            Exit Sub
        End If
        Dim heelSelectionSet As SelectionSet = selResult.Value

        Dim nMarkersFound As Integer = 0
        Dim xyOtemplate, xyHeel As LGLEGDIA1.XY
        Dim SmallHeel As Boolean = False
        Using acTr As Transaction = acCurDb.TransactionManager.StartTransaction()
            For Each idObject As ObjectId In selectionSet.GetObjectIds()
                Dim blkXMarker As BlockReference = acTr.GetObject(idObject, OpenMode.ForRead)
                Dim resbuf As ResultBuffer = blkXMarker.GetXDataForApplication("OriginID")
                Dim strXMarkerData As String = ""
                If Not IsNothing(resbuf) Then
                    For Each typeVal As TypedValue In resbuf
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        strXMarkerData = strXMarkerData & typeVal.Value
                    Next
                End If
                If strXMarkerData.Contains(sOriginXdata) Then
                    Dim position As Point3d = blkXMarker.Position
                    'If position.X <> ptLegInsertion.X Or position.Y <> ptLegInsertion.Y Then
                    '    Continue For
                    'End If
                    If position.X <> 0 Or position.Y <> 0 Then
                        nMarkersFound = nMarkersFound + 1
                        ''------PR_MakeXY(xyOtemplate, position.X, position.Y)
                        xyOtemplate.X = position.X
                        xyOtemplate.y = position.Y
                    End If
                End If
            Next
            For Each idObject As ObjectId In heelSelectionSet.GetObjectIds()
                Dim blkXMarker As BlockReference = acTr.GetObject(idObject, OpenMode.ForRead)
                Dim resbuf As ResultBuffer = blkXMarker.GetXDataForApplication("HeelID")
                Dim strHeelData As String = ""
                If Not IsNothing(resbuf) Then
                    For Each typeVal As TypedValue In resbuf
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        If typeVal.TypeCode = DxfCode.ExtendedDataInteger16 Then
                            SmallHeel = (typeVal.Value = 1)
                            Continue For
                        End If
                        strHeelData = strHeelData & typeVal.Value
                    Next
                End If
                If strHeelData.Contains(sHeelXdata) Then
                    Dim position As Point3d = blkXMarker.Position
                    If position.X <> 0 Or position.Y <> 0 Then
                        nMarkersFound = nMarkersFound + 1
                        ''------PR_MakeXY(xyHeel, position.X, position.Y)
                        xyHeel.X = position.X
                        xyHeel.y = position.Y
                        Dim rbData As ResultBuffer = blkXMarker.GetXDataForApplication("Data")
                        If Not IsNothing(rbData) Then
                            For Each typeVal As TypedValue In rbData
                                If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                    Continue For
                                End If
                                strHeelData = typeVal.Value
                                If strHeelData.Contains("1") Then
                                    SmallHeel = True
                                End If
                            Next
                        End If
                    End If
                End If
            Next
        End Using
        ''----------Close("selection", hChan)
        ''// Check if the markers have been found, otherwise exit
        If (nMarkersFound < 2) Then
            MsgBox("Missing markers for selected foot, data not found!", 16, "Heel Contracture")
            Exit Sub
        End If
        If (nMarkersFound > 2) Then
            MsgBox("Two or more drawings of the same style exist!" & Chr(10) & "Delete the extra drawing/s and try again.", 16, "Heel Contracture")
            Exit Sub
        End If

        ''// Calculate contracture points
        Dim nHeelOff As Double
        If (SmallHeel) Then
            nHeelOff = 0.25
        Else
            nHeelOff = 0.5
        End If
        Dim nInchToCM As Double = 1
        If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
            nInchToCM = 2.54
        End If
        Dim nHeelCir = xyHeel.y - xyOtemplate.y
        Dim xyPt1, xyPt2, xyPt3 As LGLEGDIA1.XY
        xyPt1.X = xyHeel.X - (nHeelOff * nInchToCM)
        xyPt1.y = xyOtemplate.y

        Dim nSeam As Double = 0.1875 * nInchToCM
        xyPt2.X = xyHeel.X
        xyPt2.y = xyOtemplate.y + nSeam + (nHeelCir / 3)

        xyPt3.X = xyHeel.X + (nHeelOff * nInchToCM)
        xyPt3.y = xyOtemplate.y

        ''// Draw contracture
        ARMDIA1.PR_SetLayer("Notes")
        Dim ptCurveColl As Point3dCollection = New Point3dCollection
        ptCurveColl.Add(New Point3d(xyPt1.X, xyPt1.y, 0))
        ptCurveColl.Add(New Point3d(xyPt2.X, xyPt2.y, 0))
        ptCurveColl.Add(New Point3d(xyPt3.X, xyPt3.y, 0))
        PR_DrawHeelContracturePoly(ptCurveColl)

        ''-----------SetData("TextHorzJust", 1) 
        ''-----------AddEntity("text", "CONTRACTURE", xyPt3.X, xyPt3.y + 0.5)
        xyPt3.y = xyPt3.y + (0.5 * nInchToCM)
        PR_DrawHeelContractureText("CONTRACTURE", xyPt3, 0.125, 0, 1)

        ''// Reset And exit
        ''------------Execute("menu", "SetLayer", Table("find", "layer", "1")) ;
        MsgBox("Contracture drawing Complete", 0, "Heel Contracture")
    End Sub
    Private Sub PR_DrawHeelContracturePoly(ByRef PointCollection As Point3dCollection)
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
                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acPoly)
                acTrans.AddNewlyCreatedDBObject(acPoly, True)
            End Using

            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Private Sub PR_DrawHeelContractureText(ByRef sText As Object, ByRef xyInsert As LGLEGDIA1.XY, ByRef nHeight As Object, ByRef nAngle As Object, ByVal nTextmode As Double)
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
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)

            '' Create a single-line text object
            Using acText As DBText = New DBText()
                acText.Position = New Point3d(xyInsert.X, xyInsert.y, 0)
                acText.Height = nHeight
                acText.TextString = sText
                acText.Rotation = nAngle
                acText.WidthFactor = 0.6
                ''acText.HorizontalMode = nTextmode
                acText.Justify = nTextmode
                ''If acText.HorizontalMode <> TextHorizontalMode.TextLeft Then
                acText.AlignmentPoint = New Point3d(xyInsert.X, xyInsert.y, 0)
                ''End If
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acText.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsert.X, xyInsert.y, 0)))
                End If
                acBlkTblRec.AppendEntity(acText)
                acTrans.AddNewlyCreatedDBObject(acText, True)
            End Using

            '' Save the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub
    ''-----------------WH_CHAP.D-----------------------------
    Public Sub PR_DrawWaistChap()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim ptCurveColl As Point3dCollection = New Point3dCollection
        Dim ptEntOpts As PromptEntityOptions = New PromptEntityOptions("Select a Leg Profile")
        ptEntOpts.AllowNone = True
        ptEntOpts.SetRejectMessage("Only SPLine or Line")
        ptEntOpts.AddAllowedClass(GetType(Spline), True)
        ptEntOpts.AddAllowedClass(GetType(Line), True)
        Dim ptEntRes As PromptEntityResult = ed.GetEntity(ptEntOpts)
        If ptEntRes.Status <> PromptStatus.OK Then
            MsgBox("User cancelled", 0, "Waist Chap")
            Exit Sub
        End If
        Dim strProfile As String = ""
        Dim nStringLength As Integer
        Dim strID As String = ""
        Dim strLeg As String = ""
        Dim sUnits As String = ""

        Dim nMarkersRequired As Integer
        Dim nMarkersFound As Integer
        Dim xyInsertion As ARMDIA1.XY
        Try
            Dim idSpline As ObjectId = ptEntRes.ObjectId
            Using tr As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acEnt As Entity = tr.GetObject(idSpline, OpenMode.ForRead)
                If (acEnt.GetType().Name.ToUpper.Equals("SPLINE")) Then
                    Dim acSpline As Spline = TryCast(acEnt, Spline)
                    Dim fitPtsCount As Integer = acSpline.NumFitPoints
                    Dim i As Integer = 0
                    While (i < fitPtsCount)
                        ptCurveColl.Add(acSpline.GetFitPointAt(i))
                        i = i + 1
                    End While
                End If
                If (acEnt.GetType().Name.ToUpper.Equals("LINE")) Then
                    Dim acLine As Line = TryCast(acEnt, Line)
                    Dim ptStart As Point3d = acLine.StartPoint
                    Dim ptEnd As Point3d = acLine.EndPoint
                    ptCurveColl.Add(ptStart)
                    ptCurveColl.Add(ptEnd)
                End If

                Dim rb As ResultBuffer = acEnt.GetXDataForApplication("ID")
                If Not IsNothing(rb) Then
                    ' Get the values in the xdata
                    For Each typeVal As TypedValue In rb
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        strProfile = strProfile & typeVal.Value
                    Next
                End If
                nStringLength = strProfile.Length
                If strProfile.Contains("LeftLegCurve") Then
                    strID = Mid(strProfile, 1, nStringLength - 8)
                    strLeg = "Left"
                ElseIf strProfile.Contains("RightLegCurve") Then
                    strID = Mid(strProfile, 1, nStringLength - 8)
                    strLeg = "Right"
                Else
                    MsgBox("A Leg Profile was not selected", 16, "Waist Chap")
                    Exit Sub
                End If

                nMarkersRequired = 1
                nMarkersFound = 0
                Dim filterChap(1) As TypedValue
                filterChap.SetValue(New TypedValue(DxfCode.BlockName, "X Marker"), 0)
                ''filterChap.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, strID + "Origin"), 1)
                filterChap.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "OriginID"), 1)
                Dim selFilterChap As SelectionFilter = New SelectionFilter(filterChap)
                Dim selResultChap As PromptSelectionResult = ed.SelectAll(selFilterChap)
                If selResultChap.Status <> PromptStatus.OK Then
                    MsgBox("Origin Marker not found", 16, "Waist Chap")
                    Exit Sub
                End If
                Dim selSetCL As SelectionSet = selResultChap.Value
                For Each idObject As ObjectId In selSetCL.GetObjectIds()
                    Dim blkXMarker As BlockReference = tr.GetObject(idObject, OpenMode.ForRead)
                    Dim rbOrigin As ResultBuffer = blkXMarker.GetXDataForApplication("OriginID")
                    Dim strOrigin As String = ""
                    If Not IsNothing(rbOrigin) Then
                        ' Get the values in the xdata
                        For Each typeVal As TypedValue In rbOrigin
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            strOrigin = strOrigin & typeVal.Value
                        Next
                    End If
                    If strOrigin.Equals(strID + "Origin") = False Then
                        Continue For
                    End If
                    Dim position As Point3d = blkXMarker.Position
                    ARMDIA1.PR_MakeXY(xyInsertion, position.X, position.Y)
                    nMarkersFound = nMarkersFound + 1
                    Dim rbUnits As ResultBuffer = blkXMarker.GetXDataForApplication("units")
                    sUnits = ""
                    If Not IsNothing(rbUnits) Then
                        ' Get the values in the xdata
                        For Each typeVal As TypedValue In rbUnits
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            sUnits = sUnits & typeVal.Value
                        Next
                    End If
                Next
                If (nMarkersFound < nMarkersRequired) Then
                    MsgBox("This is not a CHAP leg or there are missing markers!", 16, "Waist Chap")
                    Exit Sub
                End If
                If (nMarkersFound > nMarkersRequired) Then
                    MsgBox("Two or more drawings of the same style exist!" & Chr(10) & "Delete the extra drawing/s and try again.", 16, "Waist Chap")
                    Exit Sub
                End If

                Dim fileNum As Object
                fileNum = FreeFile()
                If (Not System.IO.File.Exists(fnGetSettingsPath("LookupTables") + "\\LEGCURVE.DAT")) Then
                    FileOpen(fileNum, fnGetSettingsPath("LookupTables") + "\\LEGCURVE.DAT", VB.OpenMode.Append)
                Else
                    FileOpen(fileNum, fnGetSettingsPath("LookupTables") + "\\LEGCURVE.DAT", VB.OpenMode.Output)
                End If
                ''---SetData("UnitLinearType", 0);	// "Inches"
                PrintLine(fileNum, Str(xyInsertion.X) & " " & Str(xyInsertion.y) & Chr(10))
                Dim nCount As Integer = 0
                Dim xyPt As ARMDIA1.XY
                While (nCount < ptCurveColl.Count)
                    ARMDIA1.PR_MakeXY(xyPt, ptCurveColl(nCount).X, ptCurveColl(nCount).Y)
                    PrintLine(fileNum, Str(xyPt.X) & " " & Str(xyPt.y) & Chr(10))
                    nCount = nCount + 1
                End While
                FileClose(fileNum)
                tr.Commit()
            End Using

            ''---------Draw Waist Chap Body from PR_WH_ChapBody(..) function in WHDRAW.FRM in VB folder
            Dim g_TopLegProfile, g_BotLegProfile As ARMDIA1.curve
            Dim xyOtemplate, xyChapDatum, xyProfileProximal, xyProfileDistal As ARMDIA1.XY
            xyOtemplate = xyInsertion
            Dim nInchToCM As Double = 1
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                nInchToCM = 2.54
            End If
            'Get profile points
            g_TopLegProfile.n = 0
            Dim X(ptCurveColl.Count), Y(ptCurveColl.Count) As Double
            g_TopLegProfile.X = X
            g_TopLegProfile.y = Y
            Dim nCnt As Integer = 0
            While (nCnt < ptCurveColl.Count)
                g_TopLegProfile.n = g_TopLegProfile.n + 1
                g_TopLegProfile.X(g_TopLegProfile.n) = ptCurveColl(nCnt).X
                g_TopLegProfile.y(g_TopLegProfile.n) = ptCurveColl(nCnt).Y
                nCnt = nCnt + 1
            End While

            ''Revise xyChapDatum wrt leg and last tape
            xyChapDatum.X = g_TopLegProfile.X(g_TopLegProfile.n) - (0.75 * nInchToCM)
            xyChapDatum.y = xyOtemplate.y

            ''Get the leg points that straddle x = xyChapDatum.x.
            ''This is done to take account of pleats
            Dim iStart, nn As Integer
            If g_TopLegProfile.n - 5 <= 0 Then
                iStart = 1
            Else
                iStart = g_TopLegProfile.n - 5
            End If

            ''From top
            For nn = g_TopLegProfile.n To iStart Step -1
                If g_TopLegProfile.X(nn) > xyChapDatum.X Then
                    xyProfileProximal.X = g_TopLegProfile.X(nn)
                    xyProfileProximal.Y = g_TopLegProfile.y(nn)
                End If
            Next nn

            ''From bottom
            ''NB
            '' 1. We also create the mirrored bottom leg profile
            '' 2. Special case where g_ToplegProfile.X(nn) = xyChapDatum.X
            Dim nLength As Double
            g_BotLegProfile.n = 0
            Dim X1(g_TopLegProfile.n), Y1(g_TopLegProfile.n) As Double
            g_BotLegProfile.X = X1
            g_BotLegProfile.y = Y1
            For nn = iStart To g_TopLegProfile.n
                g_BotLegProfile.n = g_BotLegProfile.n + 1
                g_BotLegProfile.X(g_BotLegProfile.n) = g_TopLegProfile.X(nn)
                nLength = g_TopLegProfile.y(nn) - xyChapDatum.y
                g_BotLegProfile.Y(g_BotLegProfile.n) = xyChapDatum.y - nLength
                If g_TopLegProfile.X(nn) <= xyChapDatum.X Then
                    If g_TopLegProfile.X(nn) = xyChapDatum.X Then
                        xyProfileDistal.X = g_TopLegProfile.X(nn) - 0.001
                    Else
                        xyProfileDistal.X = g_TopLegProfile.X(nn)
                    End If
                    xyProfileDistal.Y = g_TopLegProfile.y(nn)
                End If
            Next nn

            'Circumferences
            'nTOSCir = fnDisplayToInches(Val(txtTOSCir.Text))
            'nWaistCir = fnDisplayToInches(Val(txtWaistCir.Text))
            'nTOSHt = fnDisplayToInches(Val(txtTOSHt.Text))
            'nFoldHt = fnDisplayToInches(Val(txtFoldHt.Text))
            'nWaistHt = fnDisplayToInches(Val(txtWaistHt.Text))
            Dim nTOSCir, nWaistCir, nTOSHt, nFoldHt, nWaistHt, nWaistGivenRed, nTOSGivenRed As Double
            Dim sLegStyle As String = ""
            Using acTr As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim strLayout As String = BlockTableRecord.ModelSpace
                If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
                    strLayout = BlockTableRecord.PaperSpace
                End If
                Dim acBlkTbl As BlockTable
                acBlkTbl = acTr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                If Not acBlkTbl.Has("WAISTBODY") Then
                    MsgBox("Can't find WAISTBODY symbol to update!", 16, "Waist Chap")
                    Exit Sub
                End If
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTr.GetObject(acBlkTbl(strLayout), OpenMode.ForRead)
                For Each objID As ObjectId In acBlkTblRec
                    Dim dbObj As DBObject = acTr.GetObject(objID, OpenMode.ForRead)
                    If TypeOf dbObj Is BlockReference Then
                        Dim blkRef As BlockReference = dbObj
                        If blkRef.Name = "WAISTBODY" Then
                            For Each attributeID As ObjectId In blkRef.AttributeCollection
                                Dim attRefObj As DBObject = acTr.GetObject(attributeID, OpenMode.ForWrite)
                                If TypeOf attRefObj IsNot AttributeReference Then
                                    Continue For
                                End If
                                Dim acAttDef As AttributeReference = attRefObj
                                If acAttDef.Tag.ToUpper().Equals("TOSCir", StringComparison.InvariantCultureIgnoreCase) Then
                                    nTOSCir = fnDisplayToInches(Val(acAttDef.TextString), sUnits)
                                ElseIf acAttDef.Tag.ToUpper().Equals("WaistCir", StringComparison.InvariantCultureIgnoreCase) Then
                                    nWaistCir = fnDisplayToInches(Val(acAttDef.TextString), sUnits)
                                ElseIf acAttDef.Tag.ToUpper().Equals("TOSHt", StringComparison.InvariantCultureIgnoreCase) Then
                                    nTOSHt = fnDisplayToInches(Val(acAttDef.TextString), sUnits)
                                ElseIf acAttDef.Tag.ToUpper().Equals("FoldHt", StringComparison.InvariantCultureIgnoreCase) Then
                                    nFoldHt = fnDisplayToInches(Val(acAttDef.TextString), sUnits)
                                ElseIf acAttDef.Tag.ToUpper().Equals("WaistHt", StringComparison.InvariantCultureIgnoreCase) Then
                                    nWaistHt = fnDisplayToInches(Val(acAttDef.TextString), sUnits)
                                ElseIf acAttDef.Tag.ToUpper().Equals("WaistGivenRed", StringComparison.InvariantCultureIgnoreCase) Then
                                    nWaistGivenRed = Val(acAttDef.TextString)
                                ElseIf acAttDef.Tag.ToUpper().Equals("TOSGivenRed", StringComparison.InvariantCultureIgnoreCase) Then
                                    nTOSGivenRed = Val(acAttDef.TextString)
                                ElseIf acAttDef.Tag.ToUpper().Equals("ID", StringComparison.InvariantCultureIgnoreCase) Then
                                    sLegStyle = acAttDef.TextString
                                End If
                            Next
                        End If
                    End If
                Next
            End Using
            Dim nReduction, g_nChapCirGiven, g_nChapCirFigured, g_nChapLenGiven As Double
            If nTOSCir <> 0 And nTOSCir < nWaistCir Then
                nReduction = nTOSGivenRed
                g_nChapCirGiven = nTOSCir
                g_nChapCirFigured = (nTOSCir * nReduction) / 2
                g_nChapLenGiven = nTOSHt - nFoldHt
            Else
                nReduction = nWaistGivenRed
                g_nChapCirGiven = nWaistCir
                g_nChapCirFigured = ((nWaistCir * nReduction) / 2)
                g_nChapLenGiven = (nWaistHt - nFoldHt)
            End If
            Dim g_nChapLenFigured As Double = ((g_nChapLenGiven / 1.2) + 0.75) * nInchToCM
            Dim g_nLegStyle As Double = WHBODDIA1.fnGetNumber(sLegStyle, 1)

            ''Get intersections on leg profile
            ''Note special case for only 3 points in leg profile
            If g_TopLegProfile.n - 5 <= 0 Then
                iStart = 1
            Else
                iStart = g_TopLegProfile.n - 5
            End If
            '
            'PR_MakeXY xyOrigin, 0, xyChapDatum.y
            Dim xyOrigin As ARMDIA1.XY = xyOtemplate
            Dim xyStartCL, xyStartOFF, xyPt1, xyPt2 As ARMDIA1.XY
            ARMDIA1.PR_MakeXY(xyStartCL, xyChapDatum.X - (2.5 * nInchToCM), xyChapDatum.y)
            ARMDIA1.PR_MakeXY(xyStartOFF, xyStartCL.X, xyChapDatum.y + 100)
            ARMDIA1.PR_MakeXY(xyPt1, g_TopLegProfile.X(iStart), g_TopLegProfile.y(iStart))

            Dim Sucess As Boolean = False
            If g_TopLegProfile.n = 3 Then
                iStart = 0 ''NB 2 tape leg
            End If
            Dim ii As Integer
            For ii = iStart + 1 To g_TopLegProfile.n
                ARMDIA1.PR_MakeXY(xyPt2, g_TopLegProfile.X(ii), g_TopLegProfile.y(ii))
                If FN_LinLinInt(xyStartCL, xyStartOFF, xyPt1, xyPt2, xyStartOFF) Then
                    Sucess = True
                    Exit For
                End If
                xyPt1 = xyPt2
            Next ii

            ''assume a short 3 tape leg
            ''and extend the bottom leg profile

            If Not Sucess Then
                ARMDIA1.PR_MakeXY(xyStartOFF, xyStartCL.X, g_TopLegProfile.y(1))
                Dim TmpBotlegProfile As ARMDIA1.curve
                Dim XTmp(g_BotLegProfile.n), YTmp(g_BotLegProfile.n) As Double
                TmpBotlegProfile.X = XTmp
                TmpBotlegProfile.y = YTmp
                TmpBotlegProfile.n = g_BotLegProfile.n
                For ii = 1 To g_BotLegProfile.n
                    TmpBotlegProfile.X(ii + 1) = g_BotLegProfile.X(ii)
                    TmpBotlegProfile.Y(ii + 1) = g_BotLegProfile.y(ii)
                Next ii

                nLength = xyStartOFF.y - xyChapDatum.y
                TmpBotlegProfile.Y(1) = xyChapDatum.y - nLength
                TmpBotlegProfile.X(1) = xyStartOFF.X
                TmpBotlegProfile.n = g_BotLegProfile.n + 1
                g_BotLegProfile = TmpBotlegProfile
            End If

            ''Get intesection at X = xyChapDatum.X
            Dim xyChapDatumOFF, xyEOSCL, xyVelcroOverLapProx, xyVelcroOverLapDist, xyWaistBandProx, xyWaistBandDist, xyPt3 As ARMDIA1.XY
            Sucess = False
            ARMDIA1.PR_MakeXY(xyChapDatumOFF, xyChapDatum.X, xyChapDatum.y + 100)
            If FN_LinLinInt(xyChapDatum, xyChapDatumOFF, xyProfileDistal, xyProfileProximal, xyChapDatumOFF) = True Then
                Sucess = True
            End If
            'Give an error message here if we can't find a leg intesection
            If Not Sucess Then
                MsgBox("Can't form upper part of back edge of chap. Unable to find an intesection on the leg profile", 16, "Waist Chap")
                ''Close #fNum
                Exit Sub
            End If
            'Calculate key points
            Dim nEighthOfWaistCir As Double = (g_nChapCirFigured / 4) * nInchToCM
            ARMDIA1.PR_MakeXY(xyEOSCL, xyChapDatum.X + g_nChapLenFigured, xyChapDatum.y)
            ARMDIA1.PR_MakeXY(xyVelcroOverLapProx, xyEOSCL.X, xyChapDatum.y + nEighthOfWaistCir)
            ARMDIA1.PR_MakeXY(xyVelcroOverLapDist, xyEOSCL.X - (1 * nInchToCM), xyChapDatum.y + nEighthOfWaistCir)

            ARMDIA1.PR_MakeXY(xyWaistBandProx, xyVelcroOverLapProx.X, xyVelcroOverLapProx.y - (g_nChapCirFigured * nInchToCM))
            ARMDIA1.PR_MakeXY(xyWaistBandDist, xyVelcroOverLapDist.X, xyVelcroOverLapDist.y - (g_nChapCirFigured * nInchToCM))

            ''Draw CHAP
            ''fNum = FN_ChapOpen("C:\JOBST\DRAW.D")
            ARMDIA1.PR_SetLayer("Template" & strLeg)
            PR_DrawLine(xyWaistBandDist, xyWaistBandProx)
            PR_DrawLine(xyWaistBandProx, xyVelcroOverLapProx)

            ''Draw velcro OVERLAP
            ''If two legs and we are drawing the Left then don't draw the
            ''velcro overlap
            If g_nLegStyle = 7 And strLeg = "Left" Then
                PR_DrawLine(xyVelcroOverLapProx, xyVelcroOverLapDist)
            Else
                ''Label overlap so it can be deleted by the Zipper programme
                ARMDIA1.PR_MakeXY(xyPt1, xyVelcroOverLapProx.X, xyVelcroOverLapProx.y + (2.5 * nInchToCM))
                ARMDIA1.PR_MakeXY(xyPt2, xyVelcroOverLapDist.X, xyVelcroOverLapDist.y + (2.5 * nInchToCM))
                PR_DrawLine(xyVelcroOverLapProx, xyPt1)
                ''PR_AddDBValueToLast "Zipper", "VelcroOverlap"
                SetDBData("Zipper", "VelcroOverlap")
                PR_DrawLine(xyPt1, xyPt2)
                ''PR_AddDBValueToLast "Zipper", "VelcroOverlap"
                SetDBData("Zipper", "VelcroOverlap")
                PR_DrawLine(xyVelcroOverLapDist, xyPt2)
                ''PR_AddDBValueToLast "Zipper", "VelcroOverlap"
                SetDBData("Zipper", "VelcroOverlap")

                ARMDIA1.PR_SetLayer("Notes")
                PR_DrawLine(xyVelcroOverLapProx, xyVelcroOverLapDist)
                ''PR_AddDBValueToLast "Zipper", "VelcroOverlap"
                SetDBData("Zipper", "VelcroOverlap")
                PR_CalcMidPoint(xyPt1, xyVelcroOverLapDist, xyPt3)
                ''PR_InsertSymbol "TextAsSymbol", xyPt3, 1, 90
                PR_DrawTextAsSymbol(xyPt3, strID, "2-1/2" & Chr(34) & " VELCRO", 90)
                ''PR_AddDBValueToLast "Data", "2-1/2\"" VELCRO"
                SetDBData("Data", "2-1/2" & Chr(34) & " VELCRO")
                ''PR_AddDBValueToLast "Zipper", "VelcroOverlap"
                SetDBData("Zipper", "VelcroOverlap")
                ARMDIA1.PR_SetLayer("Template" & strLeg)
            End If

            ''Add One waist band for Chap Both Legs
            If g_nLegStyle = 7 Then
                ARMDIA1.PR_SetLayer("Notes")
                ''PR_SetTextData HORIZ_CENTER, VERT_CENTER, 0.125, CURRENT, CURRENT
                ''g_nCurrTextAngle = 90
                PR_CalcMidPoint(xyEOSCL, xyVelcroOverLapProx, xyPt3)
                xyPt3.X = xyPt3.X - 0.5
                PR_DrawText("ONE WAIST BAND", xyPt3, 0.125, 90)
                ''g_nCurrTextAngle = 0
                ARMDIA1.PR_SetLayer("Template" & strLeg)
            End If

            ''Get the first aproximation for the Top and Bottom arcs
            ''Note we work on the +ve side of the line Y = xyChapDatum.Y
            ''and then mirror the results for the bottom arc

            ''Top arc/s
            Dim nTopRadius As Double = (xyVelcroOverLapDist.X - xyChapDatum.X) / 2
            Dim xyCenTop, xyArcBottomPt, xyCenTopProx, xyCenTopDist, xyCenBot, xyCenBotProx, xyTmpProx As ARMDIA1.XY
            ARMDIA1.PR_MakeXY(xyCenTop, xyChapDatum.X + nTopRadius, xyChapDatum.y + nTopRadius)

            ''Adjust to fit if the center of the arc is above the velcro line
            ''or the end of the profile
            ARMDIA1.PR_MakeXY(xyArcBottomPt, xyCenTop.X, xyChapDatum.y)
            Dim aAngle, nBotRaduisProx As Double
            If xyCenTop.y > xyVelcroOverLapDist.y Then
                'Revise center of arc and Raduis
                ARMDIA1.PR_MakeXY(xyArcBottomPt, xyCenTop.X, xyChapDatum.y)
                nLength = ARMDIA1.FN_CalcLength(xyArcBottomPt, xyVelcroOverLapDist) / 2
                aAngle = 90 - ARMDIA1.FN_CalcAngle(xyArcBottomPt, xyVelcroOverLapDist)
                Dim nTopRaduisProx As Double = nLength / System.Math.Cos(aAngle * (ARMDIA1.PI / 180))
                ARMDIA1.PR_MakeXY(xyCenTopProx, xyArcBottomPt.X, xyArcBottomPt.y + nTopRaduisProx)
                PR_DrawArc(xyCenTopProx, xyArcBottomPt, xyVelcroOverLapDist)
            Else
                ARMDIA1.PR_MakeXY(xyPt1, xyCenTop.X + nTopRadius, xyCenTop.y * nInchToCM)
                PR_DrawLine(xyVelcroOverLapDist, xyPt1)
                PR_DrawArc(xyCenTop, xyArcBottomPt, xyPt1)
            End If

            If xyCenTop.y > xyChapDatumOFF.y Then
                'Revise center of arc and Raduis
                nLength = ARMDIA1.FN_CalcLength(xyArcBottomPt, xyChapDatumOFF) / 2
                aAngle = ARMDIA1.FN_CalcAngle(xyArcBottomPt, xyChapDatumOFF) - 90
                Dim nTopRaduisDist As Double = nLength / System.Math.Cos(aAngle * (ARMDIA1.PI / 180))
                ARMDIA1.PR_MakeXY(xyCenTopDist, xyArcBottomPt.X, xyArcBottomPt.y + nTopRaduisDist)
                PR_DrawArc(xyCenTopDist, xyChapDatumOFF, xyArcBottomPt)
            Else
                ARMDIA1.PR_MakeXY(xyPt1, xyCenTop.X - nTopRadius, xyCenTop.y)
                PR_DrawLine(xyChapDatumOFF, xyPt1)
                PR_DrawArc(xyCenTop, xyPt1, xyArcBottomPt)
            End If

            ''Bottom arc/s
            ''Calculate for +ve Y and then mirror results
            ARMDIA1.PR_MakeXY(xyArcBottomPt, xyCenTop.X, xyWaistBandDist.y + nEighthOfWaistCir)

            ''Mirror our two control points in the line Y = xyChapDatum.Y
            xyArcBottomPt.y = xyChapDatum.y + System.Math.Abs(xyArcBottomPt.y - xyChapDatum.y)
            xyWaistBandDist.y = xyChapDatum.y + System.Math.Abs(xyWaistBandDist.y - xyChapDatum.y)

            Dim nBotRadius As Double = nTopRadius * nInchToCM
            ARMDIA1.PR_MakeXY(xyCenBot, xyArcBottomPt.X, xyArcBottomPt.y + nBotRadius)
            Dim iProxCase, iDistCase As Integer
            If xyCenBot.y > xyWaistBandDist.y Then
                nLength = ARMDIA1.FN_CalcLength(xyArcBottomPt, xyWaistBandDist) / 2
                aAngle = ARMDIA1.FN_CalcAngle(xyArcBottomPt, xyWaistBandDist) - 90
                nBotRaduisProx = nLength / System.Math.Cos(aAngle * (ARMDIA1.PI / 180))
                ARMDIA1.PR_MakeXY(xyCenBotProx, xyArcBottomPt.X, xyArcBottomPt.y + nBotRaduisProx)
                xyTmpProx = xyWaistBandDist
                iProxCase = 1
            Else
                xyCenBotProx = xyCenBot
                nBotRaduisProx = nBotRadius
                ARMDIA1.PR_MakeXY(xyTmpProx, xyCenBot.X + nBotRadius, xyCenBot.y)
                iProxCase = 2
            End If

            Dim xyCenBotDist, xyTmpDist, xyChapBotDatumOFF, xyStartBotOFF As ARMDIA1.XY
            Dim nA, nB, nC, nCosTheta, aTheta As Double
            If xyCenBot.y < xyChapDatumOFF.y Then
                ARMDIA1.PR_MakeXY(xyTmpDist, xyCenBot.X - nBotRadius, xyCenBot.y)
                xyCenBotDist = xyCenBot
                iDistCase = 1
            ElseIf xyArcBottomPt.y < xyChapDatumOFF.y Then
                'Revise center of arc and Raduis
                nLength = ARMDIA1.FN_CalcLength(xyArcBottomPt, xyChapDatumOFF) / 2
                aAngle = 90 - ARMDIA1.FN_CalcAngle(xyArcBottomPt, xyChapDatumOFF)
                Dim nBotRaduisDist As Double = nLength / System.Math.Cos(aAngle * (ARMDIA1.PI / 180))
                ARMDIA1.PR_MakeXY(xyCenBotDist, xyArcBottomPt.X, xyArcBottomPt.y + nBotRaduisDist)
                iDistCase = 2
            ElseIf xyArcBottomPt.y = xyChapDatumOFF.y Then
                iDistCase = 3
            Else
                nA = ARMDIA1.FN_CalcLength(xyCenBotProx, xyChapDatumOFF)
                nC = nBotRaduisProx
                ''nB = Sqr(Abs(nA ^ 2 - nC ^ 2))
                nB = System.Math.Sqrt(System.Math.Abs(nA ^ 2 - nC ^ 2))
                'Using the cosine rule to get the angle
                nCosTheta = (nC ^ 2 - (nA ^ 2 + nB ^ 2)) / -(2 * nA * nB)
                ''aTheta = Arccos(nCosTheta) * (180 / PI)
                aTheta = System.Math.Acos(nCosTheta) * (180 / ARMDIA1.PI)
                aAngle = ARMDIA1.FN_CalcAngle(xyChapDatumOFF, xyCenBotProx) - aTheta
                ARMDIA1.PR_CalcPolar(xyChapDatumOFF, aAngle, nB, xyArcBottomPt)
                iDistCase = 4
            End If

            ''Mirror our points in the line Y = xyChapDatum.Y
            xyArcBottomPt.y = xyChapDatum.y - System.Math.Abs(xyArcBottomPt.y - xyChapDatum.y)
            xyWaistBandDist.y = xyChapDatum.y - System.Math.Abs(xyWaistBandDist.y - xyChapDatum.y)
            xyCenBotProx.y = xyChapDatum.y - System.Math.Abs(xyCenBotProx.y - xyChapDatum.y)
            xyCenBotDist.y = xyChapDatum.y - System.Math.Abs(xyCenBotDist.y - xyChapDatum.y)
            xyChapBotDatumOFF = xyChapDatumOFF
            xyChapBotDatumOFF.Y = xyChapDatum.y - System.Math.Abs(xyChapBotDatumOFF.Y - xyChapDatum.y)
            xyStartBotOFF = xyStartOFF
            xyStartBotOFF.Y = xyChapDatum.y - System.Math.Abs(xyStartBotOFF.Y - xyChapDatum.y)
            xyTmpProx.y = xyChapDatum.y - System.Math.Abs(xyTmpProx.y - xyChapDatum.y)
            xyTmpDist.y = xyChapDatum.y - System.Math.Abs(xyTmpDist.y - xyChapDatum.y)

            ''Draw the arcs and lines based on the draw case
            If iProxCase = 2 Then
                PR_DrawLine(xyWaistBandDist, xyTmpProx)
            End If
            ''PR_DrawArc(xyCenBotProx, xyArcBottomPt, xyTmpProx)
            PR_DrawArc(xyCenBotProx, xyTmpProx, xyArcBottomPt)

            Select Case iDistCase
                Case 1
                    PR_DrawLine(xyChapBotDatumOFF, xyTmpDist)
                    ''PR_DrawArc(xyCenBotDist, xyTmpDist, xyArcBottomPt)
                    PR_DrawArc(xyCenBotDist, xyArcBottomPt, xyTmpDist)
                Case 2
                    ''PR_DrawArc(xyCenBotDist, xyChapBotDatumOFF, xyArcBottomPt)
                    PR_DrawArc(xyCenBotDist, xyArcBottomPt, xyChapBotDatumOFF)
                Case 3, 4
                    PR_DrawLine(xyArcBottomPt, xyChapBotDatumOFF)
            End Select
            ''Bottom leg profile
            PR_DrawFitted(g_BotLegProfile)
            PR_DrawLine(xyStartCL, xyEOSCL)

            ''Draw markers for use with zippers
            ARMDIA1.PR_SetLayer("Construct")
            ''PR_DrawMarker xyChapDatum, "Fold"
            PR_DrawXMarker(xyChapDatum)
            SetDBData("ID", strID + "Fold")
            ''PR_DrawMarker xyEOSCL, "EOS"
            PR_DrawXMarker(xyEOSCL)
            SetDBData("ID", strID + "EOS")

            ''Draw Closing lines
            ARMDIA1.PR_SetLayer("Template" & strLeg)
            PR_DrawLine(xyStartCL, xyStartBotOFF)
            PR_DrawLine(xyOrigin, xyStartCL)

            ''Tram lines
            ARMDIA1.PR_SetLayer("Notes")
            xyPt1 = xyOrigin
            xyPt1.y = xyOrigin.y + (0.1875 * nInchToCM)
            xyPt2 = xyStartCL
            xyPt2.y = xyStartCL.y + (0.1875 * nInchToCM)
            PR_DrawLine(xyPt1, xyPt2)

            xyPt1.y = xyPt2.y + (0.5 * nInchToCM)
            xyPt2.y = xyPt2.y + (0.5 * nInchToCM)
            PR_DrawLine(xyPt1, xyPt2)

            '     PR_DrawMarker xyCenTop
            '     PR_DrawText "xyCenTop", xyCenTop, .1
            '     PR_DrawMarker xyStartOFF
            '     PR_DrawText "xyStartOFF", xyStartOFF, .1
            '     PR_DrawMarker xyChapDatumOFF
            '     PR_DrawText "xyChapDatumOFF", xyChapDatumOFF, .1
            '     PR_DrawMarker xyChapDatum
            '     PR_DrawText "xyChapDatum", xyChapDatum, .1
            '     PR_DrawMarker xyEOSCL
            '     PR_DrawText "xyEOSCL", xyEOSCL, .1
            '     PR_DrawMarker xyProfileDistal
            '     PR_DrawText "xyProfileDistal", xyProfileDistal, .1
            '     PR_DrawMarker xyProfileProximal
            '     PR_DrawText "xyProfileProximal", xyProfileProximal, .1

            ''Close #fNum
        Catch ex As Exception

        End Try
    End Sub
    Sub PR_DrawXMarker(ByRef xyInsertion As ARMDIA1.XY)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim xyStart, xyEnd, xyBase, xySecSt, xySecEnd As ARMDIA1.XY
        ARMDIA1.PR_CalcPolar(xyBase, 135, 0.0625, xyStart)
        ARMDIA1.PR_CalcPolar(xyBase, -45, 0.0625, xyEnd)
        ARMDIA1.PR_CalcPolar(xyBase, 45, 0.0625, xySecSt)
        ARMDIA1.PR_CalcPolar(xyBase, -135, 0.0625, xySecEnd)
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
                Dim acLine As Line = New Line(New Point3d(xyStart.X, xyStart.y, 0), New Point3d(xyEnd.X, xyEnd.y, 0))
                blkTblRecCross.AppendEntity(acLine)
                acLine = New Line(New Point3d(xySecSt.X, xySecSt.y, 0), New Point3d(xySecEnd.X, xySecEnd.y, 0))
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
                Dim blkRef As BlockReference = New BlockReference(New Point3d(xyInsertion.X, xyInsertion.y, 0), blkRecId)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    blkRef.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.y, 0)))
                End If
                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)
                acBlkTblRec.AppendEntity(blkRef)
                acTrans.AddNewlyCreatedDBObject(blkRef, True)
                idLastCreated = blkRef.ObjectId
            End If
            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Private Sub PR_DrawFitted(ByRef Profile As ARMDIA1.curve)
        Select Case Profile.n
            Case 0 To 1
                Exit Sub
            Case 3
                PR_DrawPoly(Profile)
            Case Else
                Dim ptColl As Point3dCollection = New Point3dCollection()
                Dim ii As Short
                For ii = 1 To Profile.n
                    ptColl.Add(New Point3d(Profile.X(ii), Profile.y(ii), 0))
                Next
                PR_DrawSpline(ptColl, "")
        End Select
    End Sub
    Private Sub PR_DrawPoly(ByRef Profile As ARMDIA1.curve)
        Dim ii As Short
        'Exit if nothing to draw
        If Profile.n <= 1 Then Exit Sub
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
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
                For ii = 1 To Profile.n
                    acPoly.AddVertexAt(ii - 1, New Point2d(Profile.X(ii), Profile.y(ii)), 0, 0, 0)
                Next ii

                'If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                '    acPoly.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.y, 0)))
                'End If
                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acPoly)
                acTrans.AddNewlyCreatedDBObject(acPoly, True)
            End Using
            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Private Sub PR_DrawArc(ByRef xyCen As ARMDIA1.XY, ByRef xyArcStart As ARMDIA1.XY, ByRef xyArcEnd As ARMDIA1.XY)
        Dim nEndAng, nRad, nStartAng, nDeltaAng As Object
        nRad = ARMDIA1.FN_CalcLength(xyCen, xyArcStart)
        nStartAng = ARMDIA1.FN_CalcAngle(xyCen, xyArcStart) * (ARMDIA1.PI / 180)
        nEndAng = ARMDIA1.FN_CalcAngle(xyCen, xyArcEnd) * (ARMDIA1.PI / 180)
        nDeltaAng = nEndAng - nStartAng

        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                                        OpenMode.ForWrite)

            Using acArc As Arc = New Arc(New Point3d(xyCen.X, xyCen.y, 0),
                                     nRad, nStartAng, nEndAng)
                'If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                '    acArc.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyCen.X, xyCen.y, 0)))
                'End If

                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acArc)
                acTrans.AddNewlyCreatedDBObject(acArc, True)
            End Using
            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Function fnDisplayToInches(ByVal nDisplay As Double, ByRef sUnits As String) As Double
        Dim iInt, iSign As Short
        Dim nDec As Double
        'retain sign
        iSign = System.Math.Sign(nDisplay)
        nDisplay = System.Math.Abs(nDisplay)
        Dim nUnitsFac As Double = 1
        If sUnits = "cm" Then
            nUnitsFac = 10 / 25.4
        End If

        'Simple case where Units are CM
        If nUnitsFac <> 1 Then
            fnDisplayToInches = WHBODDIA1.fnRoundInches(nDisplay * nUnitsFac) * iSign
            Exit Function
        End If
        'Imperial units
        iInt = Int(nDisplay)
        nDec = nDisplay - iInt

        'Check that conversion is possible (return -1 if not)
        If nDec > 0.8 Then
            fnDisplayToInches = -1
        Else
            fnDisplayToInches = WHBODDIA1.fnRoundInches(iInt + (nDec * 0.125 * 10)) * iSign
        End If
    End Function
    Function FN_LinLinInt(ByRef xyLine1Start As ARMDIA1.XY, ByRef xyLine1End As ARMDIA1.XY, ByRef xyLine2Start As ARMDIA1.XY, ByRef xyLine2End As ARMDIA1.XY, ByRef xyInt As ARMDIA1.XY) As Short
        Dim nY, nSlope1, nK2, nK1, nM1, nCase, nX As Double
        Dim nM2, nSlope2 As Object

        'Initialy false
        FN_LinLinInt = False

        'Calculate slope of lines
        nCase = 0
        nSlope1 = ARMDIA1.FN_CalcAngle(xyLine1Start, xyLine1End)
        If nSlope1 = 0 Or nSlope1 = 180 Then nCase = nCase + 1
        If nSlope1 = 90 Or nSlope1 = 270 Then nCase = nCase + 2

        nSlope2 = ARMDIA1.FN_CalcAngle(xyLine2Start, xyLine2End)
        If nSlope2 = 0 Or nSlope2 = 180 Then nCase = nCase + 4
        If nSlope2 = 90 Or nSlope2 = 270 Then nCase = nCase + 8

        Select Case nCase
            Case 0
                'Both lines are Non-Orthogonal Lines
                nM1 = (xyLine1End.y - xyLine1Start.y) / (xyLine1End.X - xyLine1Start.X) 'Slope
                nM2 = (xyLine2End.y - xyLine2Start.y) / (xyLine2End.X - xyLine2Start.X) 'Slope
                If (nM1 = nM2) Then Exit Function 'Parallel lines
                nK1 = xyLine1Start.y - (nM1 * xyLine1Start.X) 'Y-Axis intercept
                nK2 = xyLine2Start.y - (nM2 * xyLine2Start.X) 'Y-Axis intercept
                If (nK1 = nK2) Then Exit Function
                'Find X
                nX = (nK2 - nK1) / (nM1 - nM2)
                'Find Y
                nY = (nM1 * nX) + nK1
            Case 1
                'Line 1 is Horizontal or Line 2 is horizontal
                nM1 = (xyLine2End.y - xyLine2Start.y) / (xyLine2End.X - xyLine2Start.X) 'Slope
                nK1 = xyLine2Start.y - (nM1 * xyLine2Start.X) 'Y-Axis intercept
                nY = xyLine1Start.y
                'Solve for X at the given Y value
                nX = (nY - nK1) / nM1
            Case 2
                'Line 1 is Vertical or Line 2 is Vertical
                nM1 = (xyLine2End.y - xyLine2Start.y) / (xyLine2End.X - xyLine2Start.X) 'Slope
                nK1 = xyLine2Start.y - (nM1 * xyLine2Start.X) 'Y-Axis intercept
                nX = xyLine1Start.X
                'Solve for Y at the given X value
                nY = (nM1 * nX) + nK1
            Case 4
                'Line 1 is Horizontal or Line 2 is horizontal
                nM1 = (xyLine1End.y - xyLine1Start.y) / (xyLine1End.X - xyLine1Start.X) 'Slope
                nK1 = xyLine1Start.y - (nM1 * xyLine1Start.X) 'Y-Axis intercept
                nY = xyLine2Start.y
                'Solve for X at the given Y value
                nX = (nY - nK1) / nM1
            Case 5
                'Parallel orthogonal lines, no intersection possible
                Exit Function
            Case 6
                'Line1 is Vertical and the Line2 is Horizontal
                nX = xyLine1Start.X
                nY = xyLine2Start.y
            Case 8
                'Line 1 is Vertical or Line 2 is Vertical
                nM1 = (xyLine1End.y - xyLine1Start.y) / (xyLine1End.X - xyLine1Start.X) 'Slope
                nK1 = xyLine1Start.y - (nM1 * xyLine1Start.X) 'Y-Axis intercept
                nX = xyLine2Start.X
                'Solve for Y at the given X value
                nY = (nM1 * nX) + nK1
            Case 9
                'Line1 is Horizontal and the Line2 is Vertical
                nX = xyLine2Start.X
                nY = xyLine1Start.y
            Case 10
                'Parallel orthogonal lines, no intersection possible
                Exit Function
            Case Else
                Exit Function
        End Select

        'Ensure that the points X and Y are on the lines
        xyInt.X = nX
        xyInt.y = nY

        'Line 1
        If (nX < WHBODDIA1.min(xyLine1Start.X, xyLine1End.X) Or nX > max(xyLine1Start.X, xyLine1End.X)) Then Exit Function
        If (nY < WHBODDIA1.min(xyLine1Start.y, xyLine1End.y) Or nY > max(xyLine1Start.y, xyLine1End.y)) Then Exit Function

        'Line 2
        If (nX < WHBODDIA1.min(xyLine2Start.X, xyLine2End.X) Or nX > max(xyLine2Start.X, xyLine2End.X)) Then Exit Function
        If (nY < WHBODDIA1.min(xyLine2Start.y, xyLine2End.y) Or nY > max(xyLine2Start.y, xyLine2End.y)) Then Exit Function

        FN_LinLinInt = True
    End Function
    Sub PR_CalcMidPoint(ByRef xyStart As ARMDIA1.XY, ByRef xyEnd As ARMDIA1.XY, ByRef xyMid As ARMDIA1.XY)
        Static aAngle, nLength As Double

        aAngle = ARMDIA1.FN_CalcAngle(xyStart, xyEnd)
        nLength = ARMDIA1.FN_CalcLength(xyStart, xyEnd)
        If nLength = 0 Then
            xyMid = xyStart 'Avoid overflow
        Else
            ARMDIA1.PR_CalcPolar(xyStart, aAngle, nLength / 2, xyMid)
        End If
    End Sub
    Private Sub ScanLine(ByRef sLine As String, ByRef nNo As Double, ByRef sScale As Double,
                         ByRef nSpace As Double, ByRef n20Len As Double)
        nNo = FN_GetNumber(sLine, 1)
        sScale = FN_GetNumber(sLine, 2)
        nSpace = FN_GetNumber(sLine, 3)
        n20Len = FN_GetNumber(sLine, 4)
    End Sub
    Function FNRound(ByVal dVal As Double) As Double
        Return Math.Round(dVal)
    End Function
    Public Sub PR_DrawWaistLabel()
        Dim sPatient As String = "", sDiagnosis As String = "", sAge As String = "", sSEX As String = ""
        Dim workOrder As String = "", tempDate As String = "", tempEng As String = ""
        Dim sFileNo As String = "", sUnits As String = ""
        Dim obj As New BlockCreation.BlockCreation
        Dim blkId As ObjectId = New ObjectId()
        blkId = obj.LoadBlockInstance()
        If (blkId.IsNull()) Then
            MsgBox("Patient details cannot be found" & Chr(10) & "Please ensure that TITLEBOX has been used" & Chr(10) & "Then try again", 48, "Waist Label")
            Exit Sub
        End If

        obj.BindAttributes(blkId, sFileNo, sPatient, sDiagnosis, sAge, sSEX, workOrder, tempDate, tempEng, sUnits)

        Dim nAge As Double = Val(sAge)
        Dim nUnitsFac As Double = 1
        If sUnits.Contains("cm") Then
            nUnitsFac = 10 / 25.4 ''Cm To Inches 
        End If

        Dim Male As Boolean = False
        Dim Female As Boolean = False
        '' Setup SEX flags
        If (sSEX.Equals("Male")) Then
            Male = True
        Else
            Female = True
        End If

        ''-----------------------------------WHLBLBOD.D----------------------------
        Dim sLeg As String = ""
        Dim nLegStyle, nMaxCir As Double
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim sLeftThighCir As String = ""
        Dim sRightThighCir As String = ""
        Dim sCrotchStyle As String = ""
        Dim sLine As String = ""
        Dim ClosedCrotch As Boolean = False
        Dim PantyLeg As Boolean
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim blkTblRec As BlockTableRecord
            Dim blkRecId As ObjectId = ObjectId.Null
            Dim acBlkTbl As BlockTable
            Dim bIsFileNoMatch As Boolean = False
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            If Not acBlkTbl.Has("WAISTBODY") Then
                Exit Sub
            End If
            blkRecId = acBlkTbl("WAISTBODY")
            blkTblRec = acTrans.GetObject(blkRecId, OpenMode.ForRead)
            For Each objID As ObjectId In blkTblRec
                Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                If TypeOf dbObj Is AttributeDefinition Then
                    Dim acAttDef As AttributeDefinition = dbObj
                    Dim strFileNoTemp As String
                    If acAttDef.Tag.ToUpper().Equals("fileno", StringComparison.InvariantCultureIgnoreCase) Then
                        strFileNoTemp = acAttDef.TextString
                        If strFileNoTemp.Equals(sFileNo) <> True Then
                            Exit For
                        ElseIf acAttDef.Tag.ToUpper().Equals("LeftThighCir", StringComparison.InvariantCultureIgnoreCase) Then
                            sLeftThighCir = acAttDef.TextString
                        ElseIf acAttDef.Tag.ToUpper().Equals("RightThighCir", StringComparison.InvariantCultureIgnoreCase) Then
                            sRightThighCir = acAttDef.TextString
                        ElseIf acAttDef.Tag.ToUpper().Equals("CrotchStyle", StringComparison.InvariantCultureIgnoreCase) Then
                            sCrotchStyle = acAttDef.TextString
                        ElseIf acAttDef.Tag.ToUpper().Equals("ID", StringComparison.InvariantCultureIgnoreCase) Then
                            sLine = acAttDef.TextString
                        End If
                    End If
                End If
            Next
            Dim nLeftThighCir As Double = Val(sLeftThighCir)
            Dim nRightThighCir As Double = Val(sRightThighCir)
            Dim nLeftLegStyle, nRightLegStyle As Double
            ScanLine(sLine, nLegStyle, nLeftLegStyle, nRightLegStyle, nMaxCir)

            Dim LeftLeg, RightLeg As Boolean
            If ((nLeftThighCir - nRightThighCir) > 1) Then
                sLeg = "Right"
                RightLeg = True
                LeftLeg = False
            Else
                sLeg = "Left"
                LeftLeg = True
                RightLeg = False
            End If

            '' Crotch type
            Dim OpenCrotch As Boolean = False
            If sCrotchStyle.Contains("Open Crotch") Then
                OpenCrotch = True
            Else
                ClosedCrotch = True
            End If
            '' Leg  type
            If (nLegStyle = 1) Then
                PantyLeg = True
            Else
                PantyLeg = False
            End If
        End Using

        ''--------------------WHLBLMKR.D--------------------------
        Dim LeftCO_CenterArrowFound As Boolean = False
        Dim sCO_LargestTopID As String = sFileNo + sLeg + "CO_LargestTop"
        Dim sCO_LargestBottID As String = sFileNo + sLeg + "CO_LargestBott"
        Dim sCO_CenterArrowID As String = ""
        If (nLegStyle = 4) Then
            sCO_CenterArrowID = sFileNo + "Right" + "CO_CenterArrow" ''// Label Right For W4OC Right
        Else
            sCO_CenterArrowID = sFileNo + "Left" + "CO_CenterArrow" ''/ Label Left Crotch only
        End If

        Dim ed As Editor = acDoc.Editor
        '' Get the Data From X marker
        Dim filterMarker(1) As TypedValue
        filterMarker.SetValue(New TypedValue(DxfCode.BlockName, "X Marker"), 0)
        filterMarker.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "ID"), 1)
        Dim selFilterMarker As SelectionFilter = New SelectionFilter(filterMarker)
        Dim selResultMarker As PromptSelectionResult = ed.SelectAll(selFilterMarker)
        If selResultMarker.Status <> PromptStatus.OK Then
            MsgBox("Can't find Marker Details", 16, "Waist Label")
            Exit Sub
        End If
        Dim selSetCLMarker As SelectionSet = selResultMarker.Value
        Dim xyInsertionXmarker, xyCO_CenterArrow, xyCO_LargestTop, xyCO_LargestBott As ARMDIA1.XY
        Dim nOpenOff As Double
        Using tr As Transaction = acCurDb.TransactionManager.StartTransaction()
            For Each idObject As ObjectId In selSetCLMarker.GetObjectIds()
                Dim blkXMarker As BlockReference = tr.GetObject(idObject, OpenMode.ForRead)
                Dim position As Point3d = blkXMarker.Position
                ARMDIA1.PR_MakeXY(xyInsertionXmarker, position.X, position.Y)
                If xyInsertionXmarker.X = 0 Or xyInsertionXmarker.y = 0 Then
                    Continue For
                End If
                Dim resbuf As ResultBuffer = blkXMarker.GetXDataForApplication("ID")
                Dim sTmp As String = ""
                If Not IsNothing(resbuf) Then
                    ' Get the values in the xdata
                    For Each typeVal As TypedValue In resbuf
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        sTmp = sTmp & typeVal.Value
                    Next
                End If
                If sTmp.Contains(sCO_CenterArrowID) Then
                    xyCO_CenterArrow = xyInsertionXmarker
                    LeftCO_CenterArrowFound = True
                    Dim rbData As ResultBuffer = blkXMarker.GetXDataForApplication("Data")
                    Dim sData As String = ""
                    If Not IsNothing(rbData) Then
                        ' Get the values in the xdata
                        For Each typeVal As TypedValue In rbData
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            sData = sData & typeVal.Value
                        Next
                    End If
                    nOpenOff = Val(sTmp)
                ElseIf sTmp.Contains(sCO_LargestTopID) Then
                    xyCO_LargestTop = xyInsertionXmarker
                ElseIf sTmp.Contains(sCO_LargestBottID) Then
                    xyCO_LargestBott = xyInsertionXmarker
                End If
            Next
        End Using
        If (LeftCO_CenterArrowFound = False) Then
            If (nLegStyle = 4) Then
                MsgBox("RIGHT leg crotch not found to label!" & Chr(10), 16, "Waist Label")
            Else
                MsgBox("LEFT leg crotch not found to label!" & Chr(10), 16, "Waist Label")
            End If
            Exit Sub
        End If

        ''---------------------------WHLBLCRO.D------------------------------
        Dim sCrotchText As String = ""
        Dim sGussetSize As String = ""
        Dim nCrotchSize As Double = 0
        ARMDIA1.PR_SetLayer("Notes")



        ' if (nLegStyle == 2) {
        '  	SetData("TextVertJust", 16)
        ' 	SetData("TextHorzJust", 2)
        'SetData("TextAngle", 270)	
        '}
        '  else{	
        '  	SetData("TextVertJust", 16)
        ' 	SetData("TextHorzJust", 4)
        '}


        Dim nCO_ArcDiameter As Double = FNRound(ARMDIA1.FN_CalcLength(xyCO_LargestTop, xyCO_LargestBott))
        If (ClosedCrotch) Then
            If (sCrotchStyle.Contains("Horizontal Fly")) Then
                sCrotchText = "Horizontal Fly, "
                If (nAge >= 3 And nAge <= 6) Then
                    If (nCO_ArcDiameter <= 1.5) Then
                        sCrotchText = sCrotchText + "C-1"
                        nCrotchSize = 2
                    Else
                        sCrotchText = sCrotchText + "C-2"
                        nCrotchSize = 2.5
                    End If
                End If
                If (nAge >= 7 And nAge <= 11) Then
                    If (nCO_ArcDiameter <= 2) Then
                        sCrotchText = sCrotchText + "C-2"
                        nCrotchSize = 2.5
                    Else
                        sCrotchText = sCrotchText + "C-3"
                        nCrotchSize = 2.5
                    End If
                End If
                If (nAge >= 12 And nAge <= 14) Then
                    If (nCO_ArcDiameter <= 2.25) Then
                        sCrotchText = sCrotchText + "C-3"
                        nCrotchSize = 2.5
                    End If
                    If (nCO_ArcDiameter > 2.25 And nCO_ArcDiameter <= 2.5) Then
                        sCrotchText = sCrotchText + "C-4"
                        nCrotchSize = 2.5
                    End If
                    If (nCO_ArcDiameter > 2.5) Then
                        sCrotchText = sCrotchText + "C-5"
                        nCrotchSize = 3.125
                    End If
                End If
                If (nAge >= 15) Then
                    If (nCO_ArcDiameter <= 3.5) Then
                        sCrotchText = sCrotchText + "1"
                        nCrotchSize = 3.5
                    End If
                    If (nCO_ArcDiameter > 3.5 And nCO_ArcDiameter <= 4.25) Then
                        sCrotchText = sCrotchText + "2"
                        nCrotchSize = 3.75
                    End If
                    If (nCO_ArcDiameter > 4.25 And nCO_ArcDiameter <= 5) Then
                        sCrotchText = sCrotchText + "3"
                        nCrotchSize = 4
                    End If
                    If (nCO_ArcDiameter > 5) Then
                        sCrotchText = sCrotchText + "Oversize"
                        nCrotchSize = 4
                    End If
                End If
                If (sCrotchStyle.Contains("Diagonal Fly")) Then
                    sCrotchText = "Diagonal Fly, "
                    If (nAge < 10) Then
                        sCrotchText = sCrotchText + "Small"
                        nCrotchSize = 1.75
                    End If
                    If (nAge >= 10 And nAge <= 14) Then
                        sCrotchText = sCrotchText + "Medium"
                        nCrotchSize = 2.75
                    End If
                    If (nAge >= 15) Then
                        sCrotchText = sCrotchText + "Large"
                        nCrotchSize = 3.75
                    End If
                End If
            End If
            If (sCrotchStyle.Contains("Male Mesh Gusset") Or sCrotchStyle.Contains("Mesh Gusset") Or sCrotchStyle.Contains("Female Mesh Gusset")) Then
                sCrotchText = "Mesh Gusset, "
                If (sCrotchStyle.Contains("Female Mesh Gusset")) Then
                    Female = True
                    Male = False
                End If
                If (sCrotchStyle.Contains("Male Mesh Gusset")) Then
                    Female = False
                    Male = True
                End If
                If (Male) Then
                    If (nAge <= 3) Then
                        sCrotchText = sCrotchText + Format("length", 1.75)
                        nCrotchSize = 2.125
                    End If
                    If (nAge > 3 And nAge <= 14) Then
                        sCrotchText = sCrotchText + "Boy's"
                        nCrotchSize = 3.875
                    End If
                    If (nAge >= 15) Then
                        sCrotchText = sCrotchText + "Male"
                    End If
                End If
            Else
                If (nAge <= 3) Then
                    sCrotchText = sCrotchText + Format("length", 1.75)
                End If
                If (nAge > 3 And nAge <= 14) Then
                    If (nCO_ArcDiameter <= 1.75) Then
                        sCrotchText = sCrotchText + Format("length", 1)
                    End If
                    If (nCO_ArcDiameter > 1.75 And nCO_ArcDiameter <= 3.125) Then
                        sCrotchText = sCrotchText + Format("length", 1.25)
                    End If
                    If (nCO_ArcDiameter > 3.125) Then
                        sCrotchText = sCrotchText + "Regular"
                    End If
                End If
                If (nAge >= 15) Then
                    If (nMaxCir < 45) Then
                        sCrotchText = sCrotchText + "Regular"
                    Else
                        sCrotchText = sCrotchText + "Oversize"
                    End If
                End If
            End If
            If (String.Compare(sCrotchStyle, "Gusset") = 0) Then
                sCrotchText = sCrotchStyle + ", "
                If (Male) Then
                    If (nAge <= 3) Then
                        sCrotchText = sCrotchText + Format("length", 1.75)
                        nCrotchSize = 2.125
                    End If
                    If (nAge > 3 And nAge <= 14) Then
                        sCrotchText = sCrotchText + "Boy's"
                        nCrotchSize = 3.875
                    End If
                    If (nAge >= 15) Then
                        sCrotchText = sCrotchText + "Male"
                    End If
                Else
                    If (nAge <= 3) Then
                        sCrotchText = sCrotchText + Format("length", 1.75)
                    End If
                    If (nAge > 3 And nAge <= 14) Then
                        If (nCO_ArcDiameter <= 1.75) Then
                            sCrotchText = sCrotchText + Format("length", 1)
                        End If
                        If (nCO_ArcDiameter > 1.75 And nCO_ArcDiameter <= 3.125) Then
                            sCrotchText = sCrotchText + Format("length", 1.25)
                        End If
                        If (nCO_ArcDiameter > 3.125) Then
                            sCrotchText = sCrotchText + "Regular"
                        End If
                    End If
                    If (nAge >= 15) Then
                        If (nMaxCir < 45) Then
                            sCrotchText = sCrotchText + "Regular"
                        Else
                            sCrotchText = sCrotchText + "Oversize"
                        End If
                    End If
                End If
            End If
            PR_DrawClosedArrow(xyCO_CenterArrow, 180)
            Dim xyText As ARMDIA1.XY
            ARMDIA1.PR_MakeXY(xyText, xyCO_CenterArrow.X - 0.75, xyCO_CenterArrow.y)
            PR_DrawText(sCrotchText, xyText, 0.125, 0)
        Else
            'Open Crotch
            If (nOpenOff = 0.375 Or (nAge <= 10 And PantyLeg)) Then
                ''AddEntity("text", "1/2\" & "Elastic", xyCO_CenterArrow.X - 0.75, xyCO_CenterArrow.y)
                Dim xyText As ARMDIA1.XY
                ARMDIA1.PR_MakeXY(xyText, xyCO_CenterArrow.X - 0.75, xyCO_CenterArrow.y)
                sCrotchText = "1/2" & Chr(34) & "Elastic"
                PR_DrawText(sCrotchText, xyText, 0.125, 0)
            End If
        End If
    End Sub
    Sub PR_DrawClosedArrow(ByRef xyPoint As ARMDIA1.XY, ByRef nAngle As Object)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        Dim xyStart As ARMDIA1.XY
        ''PR_MakeXY(xyStart, xyPoint.X, xyPoint.Y + 0.125)
        xyStart = xyPoint
        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)

            '' Create a polyline with two segments (3 points)
            Dim acPoly As Polyline = New Polyline()
            acPoly.AddVertexAt(0, New Point2d(xyStart.X, xyStart.y), 0, 0, 0)
            acPoly.AddVertexAt(0, New Point2d(xyStart.X + 0.125, xyStart.y + 0.0625), 0, 0, 0)
            acPoly.AddVertexAt(0, New Point2d(xyStart.X + 0.125, xyStart.y - 0.0625), 0, 0, 0)
            acPoly.AddVertexAt(0, New Point2d(xyStart.X, xyStart.y), 0, 0, 0)
            acPoly.TransformBy(Matrix3d.Rotation((nAngle * (ARMDIA1.PI / 180)), Vector3d.ZAxis, New Point3d(xyStart.X, xyStart.y, 0)))
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                acPoly.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyStart.X, xyStart.y, 0)))
            End If

            '' Add the new object to the block table record and the transaction
            Dim idPolyline As ObjectId = New ObjectId
            idPolyline = acBlkTblRec.AppendEntity(acPoly)
            acTrans.AddNewlyCreatedDBObject(acPoly, True)
            idLastCreated = acPoly.ObjectId
            ''Create Hatch entity
            Dim ObjIds As ObjectIdCollection = New ObjectIdCollection
            ObjIds.Add(idPolyline)
            Dim oHatch As Hatch = New Hatch()
            Dim normal As Vector3d = New Vector3d(0.0, 0.0, 1.0)
            oHatch.Normal = normal
            oHatch.Elevation = 0.0
            oHatch.PatternScale = 2.0
            oHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID")
            oHatch.ColorIndex = 1
            acBlkTblRec.AppendEntity(oHatch)
            acTrans.AddNewlyCreatedDBObject(oHatch, True)
            oHatch.Associative = True
            oHatch.AppendLoop(CInt(HatchLoopTypes.Default), ObjIds)
            oHatch.Color = acPoly.Color
            oHatch.EvaluateHatch(True)

            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_DrawText(ByRef sText As String, ByRef xyInsert As ARMDIA1.XY, ByRef nHeight As Object, ByRef nAngle As Double)
        Dim nWidth As Object
        nWidth = nHeight * g_nWebCurrTextAspect
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                     OpenMode.ForRead)
            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                                        OpenMode.ForWrite)
            '' Create a single-line text object
            Using acText As DBText = New DBText()
                acText.Position = New Point3d(xyInsert.X, xyInsert.y, 0)
                acText.Height = nHeight
                acText.TextString = sText
                acText.Rotation = nAngle
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acText.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsert.X, xyInsert.y, 0)))
                End If
                acBlkTblRec.AppendEntity(acText)
                acTrans.AddNewlyCreatedDBObject(acText, True)
            End Using

            '' Save the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub
    Public Sub PR_DrawKneeContracture()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim ptEntOpts As PromptEntityOptions = New PromptEntityOptions(Chr(10) & "Select a Leg Profile")
        ptEntOpts.AllowNone = True
        ''ptEntOpts.Keywords.Add("Spline")
        ptEntOpts.SetRejectMessage(Chr(10) & "Only Curve")
        ptEntOpts.AddAllowedClass(GetType(Spline), True)
        Dim ptEntRes As PromptEntityResult = ed.GetEntity(ptEntOpts)
        If ptEntRes.Status <> PromptStatus.OK Then
            MsgBox("Profile not selected", 16, "Knee Contracture")
            Exit Sub
        End If
        Dim strProfile As String = ""
        Dim strXMarkerHandle As String = ""
        Dim ptCurveColl As New Point3dCollection
        Dim fitPtsCount As Integer = 0
        Try
            Dim idSpline As ObjectId = ptEntRes.ObjectId
            Using tr As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acSpline As Spline = tr.GetObject(idSpline, OpenMode.ForRead)
                Dim rb As ResultBuffer = acSpline.GetXDataForApplication("ID")
                If Not IsNothing(rb) Then
                    ' Get the values in the xdata
                    For Each typeVal As TypedValue In rb
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        strProfile = strProfile & typeVal.Value
                    Next
                End If
                Dim rbHandle As ResultBuffer = acSpline.GetXDataForApplication("Handle")
                If Not IsNothing(rbHandle) Then
                    ' Get the values in the xdata
                    For Each typeVal As TypedValue In rbHandle
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        strXMarkerHandle = strXMarkerHandle & typeVal.Value
                    Next
                End If
                fitPtsCount = acSpline.NumFitPoints
                Dim i As Integer = 0
                While (i < fitPtsCount)
                    ptCurveColl.Add(acSpline.GetFitPointAt(i))
                    i = i + 1
                End While
            End Using
        Catch ex As Exception

        End Try
        strProfile = Trim(strProfile)
        Dim sStyleID, sLeg As String
        If strProfile.Contains("LeftLegCurve") Then
            sStyleID = Mid(strProfile, 1, strProfile.Length - 12)
            sLeg = "Left"
        ElseIf strProfile.Contains("RightLegCurve") Then
            sStyleID = Mid(strProfile, 1, strProfile.Length - 13)
            sLeg = "Right"
        Else
            MsgBox("A Leg Profile was not selected", 16, "Knee Contracture")
            Exit Sub
        End If

        ''// Get Marker data
        Dim sOtemplateID As String = sStyleID + sLeg + "Origin"
        Dim filterType(1) As TypedValue
        filterType.SetValue(New TypedValue(DxfCode.BlockName, "X Marker"), 0)
        filterType.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "OriginID"), 1)
        Dim selFilter As SelectionFilter = New SelectionFilter(filterType)
        Dim selResult As PromptSelectionResult = ed.SelectAll(selFilter)
        If selResult.Status <> PromptStatus.OK Then
            MsgBox("Origin Marker not found", 16, "Knee Contracture")
            Exit Sub
        End If
        Dim selectionSet As SelectionSet = selResult.Value
        Dim xyOtemplate As ARMDIA1.XY
        Dim OriginFound As Boolean = False
        Using acTr As Transaction = acCurDb.TransactionManager.StartTransaction()
            For Each idObject As ObjectId In selectionSet.GetObjectIds()
                Dim blkXMarker As BlockReference = acTr.GetObject(idObject, OpenMode.ForRead)
                Dim rbOrigin As ResultBuffer = blkXMarker.GetXDataForApplication("OriginID")
                Dim strOrigin As String = ""
                If Not IsNothing(rbOrigin) Then
                    ' Get the values in the xdata
                    For Each typeVal As TypedValue In rbOrigin
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        strOrigin = strOrigin & typeVal.Value
                    Next
                End If
                If strOrigin.Equals(sOtemplateID) = False Then
                    Continue For
                End If
                If strXMarkerHandle <> "" And strXMarkerHandle.Contains(blkXMarker.Handle.ToString) = False Then
                    Continue For
                End If
                Dim position As Point3d = blkXMarker.Position
                If position.X <> 0 Or position.Y <> 0 Then
                    OriginFound = True
                    ''PR_MakeXY(xyOtemplate, position.X, position.Y)
                    xyOtemplate.X = position.X
                    xyOtemplate.y = position.Y
                    'If xyOtemplate.X = LEGRIGHT.xyRightLegInsertion.X And xyOtemplate.y = LEGRIGHT.xyRightLegInsertion.y Then
                    '    Exit For
                    'End If
                End If
            Next
        End Using

        ''// Check if that the markers have been found, otherwise exit
        If (OriginFound = False) Then
            MsgBox("Origin Marker not found!", 16, "Knee Contracture")
            Exit Sub
        End If
        ''// Start dialog wrt contractures
        Dim oKneeCtrFrm As New KNEECTR_frm
        oKneeCtrFrm.g_sCaption = sLeg + " Knee Contracture"
        oKneeCtrFrm.ShowDialog()
        If oKneeCtrFrm.g_bIsCancel = True Then
            Exit Sub
        End If
        Dim nContracture As Double = 0
        Dim sTmp As String = Mid(oKneeCtrFrm.g_sContracture, 1, 2)
        If (sTmp.Equals("10")) Then nContracture = 0.5
        If (sTmp.Equals("36")) Then nContracture = 1
        If (sTmp.Equals("71")) Then nContracture = 1.5
        Dim nInchToCM As Double = 1
        If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
            nInchToCM = 2.54
        End If
        nContracture = nContracture * nInchToCM
        Dim xyKneeDip As ARMDIA1.XY
        While (True)
            Dim kneeEntOpts As PromptEntityOptions = New PromptEntityOptions(Chr(10) & "Select the Knee Dip tape symbol")
            kneeEntOpts.AllowNone = True
            kneeEntOpts.SetRejectMessage(Chr(10) & "Only block")
            kneeEntOpts.AddAllowedClass(GetType(BlockReference), True)
            Dim kneeEntRes As PromptEntityResult = ed.GetEntity(kneeEntOpts)
            If kneeEntRes.Status <> PromptStatus.OK Then
                MsgBox("Knee Dip tape symbol was not selected", 16, "Knee Contracture")
                Exit Sub
            End If
            Dim idObject As ObjectId = kneeEntRes.ObjectId
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim blkXMarker As BlockReference = acTrans.GetObject(idObject, OpenMode.ForRead)
                Dim strBlkName As String = blkXMarker.Name
                If strBlkName.Contains("tape") = False Then
                    Continue While
                End If
                Dim position As Point3d = blkXMarker.Position
                ''PR_MakeXY(xyKneeDip, position.X, position.Y)
                xyKneeDip.X = position.X
                xyKneeDip.y = position.Y
                Exit While
            End Using
        End While
        Dim nn As Integer = 1
        While (nn <= fitPtsCount)
            Dim xyTmp As ARMDIA1.XY
            ''PR_MakeXY(xyTmp, ptCurveColl(nn).X, ptCurveColl(nn).Y)
            xyTmp.X = ptCurveColl(nn).X
            xyTmp.y = ptCurveColl(nn).Y
            Dim nTolerance As Double = System.Math.Abs(xyTmp.X - xyKneeDip.X)
            If (nTolerance < 0.1) Then
                xyKneeDip = xyTmp
                Exit While
            End If
            nn = nn + 1
        End While
        If (nn = fitPtsCount + 1) Then
            MsgBox("Can't find Knee Dip vertex on selected profile.", 16, "Knee Contracture")
            Exit Sub
        End If
        ''// Get next vertex to allow construction of contracture.
        ''// Check availability of vertices to avoid nasty failure And error messages.
        ''// Note:-
        ''//       Special case where contracture width Is longer than the distance between the
        ''//       initially selected next vertex And the knee dip.
        If (nn + 1 > fitPtsCount) Then
            MsgBox("Can't calculate Knee contracture on selected profile.", 16, "Knee Contracture")
            Exit Sub
        End If
        Dim xyPt1, xyPt2, xyInt As ARMDIA1.XY
        ARMDIA1.PR_MakeXY(xyPt1, ptCurveColl(nn + 1).X, ptCurveColl(nn + 1).Y)
        xyPt1.X = ptCurveColl(nn + 1).X
        xyPt1.y = ptCurveColl(nn + 1).Y
        Dim nError As Short = FN_CirLinInt(xyKneeDip, xyPt1, xyKneeDip, nContracture, xyInt)
        If (nError = False) Then
            If (nn + 1 > fitPtsCount) Then
                MsgBox("Can't calculate Knee contracture on selected profile.", 16, "Knee Contracture")
                Exit Sub
            End If
            ARMDIA1.PR_MakeXY(xyPt2, ptCurveColl(nn + 2).X, ptCurveColl(nn + 2).Y)
            nError = FN_CirLinInt(xyPt1, xyPt2, xyKneeDip, nContracture, xyInt)
            If (nError = False) Then
                MsgBox("Can't calculate Knee contracture on selected profile.", 16, "Knee Contracture")
                Exit Sub
            End If
        End If
        xyPt2 = xyInt
        Dim nKneeCir, aAngle, nA As Double
        nKneeCir = xyKneeDip.y - xyOtemplate.y
        aAngle = ARMDIA1.FN_CalcAngle(xyKneeDip, xyPt2)
        nA = System.Math.Sqrt((0.5 * nKneeCir) ^ 2 - (0.5 * nContracture) ^ 2)
        ''----------xyPt1 = CalcXY("relpolar", CalcXY("relpolar", xyKneeDip, nContracture * 0.5, aAngle), nA, aAngle + 270)
        ARMDIA1.PR_CalcPolar(xyKneeDip, aAngle, nContracture * 0.5, xyInt)
        ARMDIA1.PR_CalcPolar(xyInt, aAngle + 270, nA, xyPt1)

        ''// Draw contracture
        ''//
        ARMDIA1.PR_SetLayer("Notes")
        'StartPoly("polyline")
        'AddVertex(xyKneeDip)
        'AddVertex(xyPt1)
        'AddVertex(xyPt2)
        'EndPoly()
        Dim ptPolyColl As Point3dCollection = New Point3dCollection
        ptPolyColl.Add(New Point3d(xyKneeDip.X, xyKneeDip.y, 0))
        ptPolyColl.Add(New Point3d(xyPt1.X, xyPt1.y, 0))
        ptPolyColl.Add(New Point3d(xyPt2.X, xyPt2.y, 0))
        PR_DrawHeelContracturePoly(ptPolyColl)

        ''------SetData("TextHorzJust", 4) ;
        ''------AddEntity("text", "CONTRACTURE", xyKneeDip.X, xyKneeDip.y - ((xyKneeDip.y - xyPt1.y) / 2));
        xyKneeDip.y = xyKneeDip.y - ((xyKneeDip.y - xyPt1.y) / 2)
        Dim xyInsert As LGLEGDIA1.XY
        xyInsert.X = xyKneeDip.X
        xyInsert.y = xyKneeDip.y
        PR_DrawHeelContractureText("CONTRACTURE", xyInsert, 0.125, 0, 3)

        ''// Display Contracture size on layer construct (for info only)
        ARMDIA1.PR_SetLayer("Construct")
        ''-------SetData("TextHorzJust", 2) ;
        ''-------AddEntity("text", sContracture, xyPt1.X, xyPt1.y - 0.25); 
        xyPt1.y = xyPt1.y - (0.25 * nInchToCM)
        xyInsert.X = xyPt1.X
        xyInsert.y = xyPt1.y
        PR_DrawHeelContractureText(oKneeCtrFrm.g_sContracture, xyInsert, 0.125, 0, 2)

    End Sub
    Function FN_CirLinInt(ByRef xyStart As ARMDIA1.XY, ByRef xyEnd As ARMDIA1.XY, ByRef xyCen As ARMDIA1.XY, ByRef nRad As Double, ByRef xyInt As ARMDIA1.XY) As Short
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

        nSlope = ARMDIA1.FN_CalcAngle(xyStart, xyEnd)

        'Horizontal Line
        If nSlope = 0 Or nSlope = 180 Then
            nSlope = -1
            nC = nRad ^ 2 - (xyStart.y - xyCen.y) ^ 2
            If nC < 0 Then
                FN_CirLinInt = False 'no roots
                Exit Function
            End If
            nSign = 1 'test each root
            While nSign > -2
                nRoot = xyCen.X + System.Math.Sqrt(nC) * nSign
                If nRoot >= MANGLOVE1.min(xyStart.X, xyEnd.X) And nRoot <= max(xyStart.X, xyEnd.X) Then
                    xyInt.X = nRoot
                    xyInt.y = xyStart.y
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
                nRoot = xyCen.y + System.Math.Sqrt(nC) * nSign
                If nRoot >= MANGLOVE1.min(xyStart.y, xyEnd.y) And nRoot <= max(xyStart.y, xyEnd.y) Then
                    xyInt.y = nRoot
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
            nM = (xyEnd.y - xyStart.y) / (xyEnd.X - xyStart.X) 'Slope
            nK = xyStart.y - nM * xyStart.X 'Y-Axis intercept
            nA = (1 + nM ^ 2)
            nB = 2 * (-xyCen.X + (nM * nK) - (xyCen.y * nM))
            nC = (xyCen.X ^ 2) + (nK ^ 2) + (xyCen.y ^ 2) - (2 * xyCen.y * nK) - (nRad ^ 2)
            nCalcTmp = (nB ^ 2) - (4 * nC * nA)

            If (nCalcTmp < 0) Then
                FN_CirLinInt = False 'No Roots
                Exit Function
            End If
            nSign = 1
            While nSign > -2
                nRoot = (-nB + (System.Math.Sqrt(nCalcTmp) / nSign)) / (2 * nA)
                If nRoot >= MANGLOVE1.min(xyStart.X, xyEnd.X) And nRoot <= max(xyStart.X, xyEnd.X) Then
                    xyInt.X = nRoot
                    xyInt.y = nM * nRoot + nK
                    FN_CirLinInt = True
                    Exit Function 'Return first root found
                End If
                nSign = nSign - 2
            End While
            FN_CirLinInt = False 'Should never get to here
        End If
        FN_CirLinInt = False
    End Function

    Public Function PR_DrawWaistBodyZipper() As Boolean
        Dim sProfileID As String = ""
        Dim sStyleID As String = ""
        Dim sLeg As String = ""
        Dim sOtemplateID As String = ""
        Dim sCO_WaistBottID As String = ""
        Dim xyOtemplate As ARMDIA1.XY
        Dim xyCO_WaistBott As ARMDIA1.XY
        Dim xyMin As ARMDIA1.XY
        Dim xyMax As ARMDIA1.XY
        Dim xyHorizEnd As ARMDIA1.XY
        Dim xyHorizStart As ARMDIA1.XY

        Dim nStringLength As Integer = 0
        Dim nMarkersFound As Integer = 0
        Dim nType As Integer = 0
        Dim nn As Integer = 0
        Dim sTmp As String = ""
        Dim sUnits As String = ""
        Dim sType As String = ""
        Dim EOStoSelectedPoint As Boolean = False

        Dim sDlgLengthList As String = ""
        Dim sDlgElasticList As String = ""
        Dim nMedial As Integer = 0
        Dim bIsLoop As Boolean = False
        Dim nMinZipLength As Integer = 0
        Dim sZipLength As String = ""
        Dim nZipLength As Integer = 0


        Dim nThighFiguredCir As Integer = 0
        Dim nWaistActualCir As Integer = 0
        Dim EOSFound As Boolean = False
        Dim nElasticProximal As Double = 0
        Dim nElastic As Double = 0
        Dim sElasticProximal As String = ""
        Dim nSeam As Double = 0
        Dim nDrawnLength As Double = 0
        Dim sZipOffset As String = ""

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim ptCurveColl As Point3dCollection = New Point3dCollection
        Dim ptEntOpts As PromptEntityOptions = New PromptEntityOptions("Select a Leg Profile")
        ptEntOpts.AllowNone = True
        ptEntOpts.SetRejectMessage("Select SPLine")
        ptEntOpts.AddAllowedClass(GetType(Spline), True)
        Dim ptEntRes As PromptEntityResult = ed.GetEntity(ptEntOpts)
        If ptEntRes.Status <> PromptStatus.OK Then
            MsgBox("User cancelled", 0, "Zipper Body")
            Return False
        End If

        Using tr As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim idObj As ObjectId = ptEntRes.ObjectId
            Dim acEnt As Entity = tr.GetObject(idObj, OpenMode.ForRead)
            Dim rb As ResultBuffer = acEnt.GetXDataForApplication("ID")
            If Not IsNothing(rb) Then
                ' Get the values in the xdata
                For Each typeVal As TypedValue In rb
                    If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                        Continue For
                    End If
                    sProfileID = sProfileID & typeVal.Value
                Next
            End If
            Dim strXMarkerHandle As String = ""
            Dim rbHandle As ResultBuffer = acEnt.GetXDataForApplication("Handle")
            If Not IsNothing(rbHandle) Then
                ' Get the values in the xdata
                For Each typeVal As TypedValue In rbHandle
                    If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                        Continue For
                    End If
                    strXMarkerHandle = strXMarkerHandle & typeVal.Value
                Next
            End If

            nStringLength = sProfileID.Length
            If String.Compare("LeftLegCurve", Mid(sProfileID, nStringLength - 11, 12)) = 0 Then
                ''sStyleID = Mid(sProfileID, 1, nStringLength - 8)
                sStyleID = Mid(sProfileID, 1, nStringLength - 12)
                sLeg = "Left"
            End If

            If String.Compare("RightLegCurve", Mid(sProfileID, nStringLength - 12, 13)) = 0 Then
                ''sStyleID = Mid(sProfileID, 1, nStringLength - 8)
                sStyleID = Mid(sProfileID, 1, nStringLength - 13)
                sLeg = "Right"
            End If

            If (sLeg.Length = 0) Then
                MsgBox("A Leg Profile was not selected", 16, "Zipper Body")
                Return False
            End If

            sOtemplateID = sStyleID + sLeg + "Origin"

            ''sStyleID = Mid(sStyleID, 5, sStyleID.Length - 4)
            sCO_WaistBottID = sStyleID + sLeg + "CO_WaistBott"

            nMarkersFound = 0

            '' Get the Data From X marker
            Dim filterMarker(1) As TypedValue
            filterMarker.SetValue(New TypedValue(DxfCode.BlockName, "X Marker"), 0)
            filterMarker.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "OriginID"), 1)
            Dim selFilterMarker As SelectionFilter = New SelectionFilter(filterMarker)
            Dim selResultMarker As PromptSelectionResult = ed.SelectAll(selFilterMarker)
            If selResultMarker.Status <> PromptStatus.OK Then
                MsgBox("Can't find Marker Details", 16, "Zipper Body")
                Return False
            End If
            Dim selSetCLMarker As SelectionSet = selResultMarker.Value
            Dim xyInsertionXmarker As ARMDIA1.XY


            For Each idObject As ObjectId In selSetCLMarker.GetObjectIds()
                Dim blkXMarker As BlockReference = tr.GetObject(idObject, OpenMode.ForRead)
                Dim position As Point3d = blkXMarker.Position
                ARMDIA1.PR_MakeXY(xyInsertionXmarker, position.X, position.Y)
                If xyInsertionXmarker.X = 0 And xyInsertionXmarker.y = 0 Then
                    Continue For
                End If
                Dim resbuf As ResultBuffer = blkXMarker.GetXDataForApplication("OriginID")
                sTmp = ""
                If Not IsNothing(resbuf) Then
                    ' Get the values in the xdata
                    For Each typeVal As TypedValue In resbuf
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        sTmp = sTmp & typeVal.Value
                    Next
                End If

                If (String.Compare(sTmp, sOtemplateID) = 0) Then
                    nMarkersFound = nMarkersFound + 1
                    xyOtemplate = xyInsertionXmarker
                    Dim strHandle As String = ""
                    Dim resbufHandle As ResultBuffer = blkXMarker.GetXDataForApplication("Handle")
                    If Not IsNothing(resbufHandle) Then
                        ' Get the values in the xdata
                        For Each typeVal As TypedValue In resbufHandle
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            strHandle = strHandle & typeVal.Value
                        Next
                    End If
                    If strHandle.Contains(strXMarkerHandle) = False Then
                        Continue For
                    End If

                    Dim rbUnits As ResultBuffer = blkXMarker.GetXDataForApplication("units")
                    If Not IsNothing(rbUnits) Then
                        ' Get the values in the xdata
                        For Each typeVal As TypedValue In rbUnits
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            sUnits = sUnits & typeVal.Value
                        Next
                    End If
                    Dim rbData As ResultBuffer = blkXMarker.GetXDataForApplication("Data")
                    sTmp = ""
                    If Not IsNothing(rbData) Then
                        ' Get the values in the xdata
                        For Each typeVal As TypedValue In rbData
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            sTmp = sTmp & typeVal.Value
                        Next
                    End If

                    If (sTmp.Equals("")) Then
                        MsgBox("Can't get data from Origin Marker!", 16, "Zipper Body")
                        Return False
                    End If
                    ScanLine(sTmp, nType, nn)
                    sType = sTmp
                End If
                If sTmp.Contains(sCO_WaistBottID) Then
                    xyCO_WaistBott = xyInsertionXmarker
                End If
            Next

            'Check If that the markers have been found, otherwise exit
            If (String.Compare("CHAP", sType) = 0) Then
                MsgBox("Can't use this zipper for a CHAP style", 16, "Zipper Body")
                Return False
            End If

            If (nMarkersFound < 1) Then
                MsgBox("Missing marker for selected Waist Height, data not found", 16, "Zipper Body")
                Return False
            End If
            If (nMarkersFound > 1) Then
                MsgBox("Two or more drawings of the same style exist!" & Chr(10) & "Delete the extra drawing/s and try again.", 16, "Zipper Body")
                Return False
            End If
            ' Create Dialog
            'Get Zipper style

            sUnits = "Inches"

            EOStoSelectedPoint = False
            '     sDlgElasticList = "3/8\" Elastic\n3/4\" Elastic\n1½\" Elastic\nNo Elastic";
            'sDlgElasticList =  "3/4\" Elastic\n" + sDlgElasticList
            'sDlgLengthList = "Give a length\nSelected Point"
            nMedial = 0
            bIsLoop = True

            nMinZipLength = 5

            While (bIsLoop)
                '           nButX = 65; nButY = 40;
                '           hDlg = Open("dialog", sLeg + " Body Zipper (Waist Height)", "font Helv 8", 20, 20, 210, 75);

                'AddControl(hDlg, "pushbutton", nButX, nButY, 35, 14, "Cancel", "%cancel", "");
                'AddControl(hDlg, "pushbutton", nButX + 48, nButY, 35, 14, "OK", "%ok", "");

                'AddControl(hDlg, "ltext", 5, 12, 28, 14, "EOS to", "string", "");
                'AddControl(hDlg, "combobox", 35, 10, 70, 40, sDlgLengthList, "string", "sZipLength");
                'AddControl(hDlg, "ltext", 110, 12, 30, 14, "Proximal:", "string", "");
                '	AddControl(hDlg, "dropdownlist", 145, 10, 60, 70, sDlgElasticList, "string", "sElasticProximal");

                '     	Ok = Display("dialog", hDlg, "%center");
                ' 	Close("dialog", hDlg);

                'If (Ok == %cancel ) Exit (%ok, "User Cancel!") ;
                Dim oZipperBodyFrm As New ZIPBOD_frm
                oZipperBodyFrm.g_sCaption = sLeg + " Body Zipper (Waist Height)"
                oZipperBodyFrm.ShowDialog()
                If oZipperBodyFrm.g_bIsCancel = True Then
                    MsgBox("User Cancel!", 16, "Zipper Body")
                    Return False
                End If
                sZipLength = oZipperBodyFrm.g_sZipLength
                sElasticProximal = oZipperBodyFrm.g_sElasticProximal

                If (String.Compare("Selected Point", sZipLength) = 0) Then
                    EOStoSelectedPoint = True
                End If

                If (EOStoSelectedPoint) Then
                    bIsLoop = False

                Else
                    nZipLength = Val(sZipLength)
                    If (nZipLength = 0 And sZipLength.Length > 0) Then
                        MsgBox("Invalid given length!", 16, "Zipper Body")
                        bIsLoop = True
                    Else
                        bIsLoop = False
                    End If
                End If
            End While
            EOSFound = False
            Dim filterLine(2) As TypedValue
            filterLine.SetValue(New TypedValue(DxfCode.Start, "Line"), 0)
            filterLine.SetValue(New TypedValue(DxfCode.LayerName, "Construct"), 1)
            ''filterLine.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, sStyleID + sLeg + "EOSLine"), 2)
            filterLine.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "EOSID"), 2)
            Dim selFilterLine As SelectionFilter = New SelectionFilter(filterLine)
            Dim selResultLine As PromptSelectionResult = ed.SelectAll(selFilterLine)
            If selResultLine.Status <> PromptStatus.OK Then
                MsgBox("Can't find an EOS for End of Zipper at the PROXIMAL EOS!", 16, "Zipper Body")
                Return False
            End If
            Dim selSetLine As SelectionSet = selResultLine.Value
            Dim extLine As Extents3d

            For Each idLine As ObjectId In selSetLine.GetObjectIds()
                Dim line As Line = tr.GetObject(idLine, OpenMode.ForRead)
                sTmp = ""
                Dim rbEOS As ResultBuffer = line.GetXDataForApplication("EOSID")
                If Not IsNothing(rbEOS) Then
                    ' Get the values in the xdata
                    For Each typeVal As TypedValue In rbEOS
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        sTmp = sTmp & typeVal.Value
                    Next
                End If
                If sTmp.Contains(sStyleID + sLeg + "EOSLine") = False Then
                    Continue For
                End If
                extLine = line.GeometricExtents
                Dim ptMin As Point3d = extLine.MinPoint
                Dim ptMax As Point3d = extLine.MaxPoint
                ARMDIA1.PR_MakeXY(xyMin, ptMin.X, ptMin.Y)
                ARMDIA1.PR_MakeXY(xyMax, ptMax.X, ptMax.Y)
                xyHorizEnd.X = xyMin.X
                sTmp = ""
                Dim rbLineData As ResultBuffer = line.GetXDataForApplication("Data")
                If Not IsNothing(rbLineData) Then
                    ' Get the values in the xdata
                    For Each typeVal As TypedValue In rbLineData
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        sTmp = sTmp & typeVal.Value
                    Next
                End If

                If (sTmp.Equals("")) Then
                    MsgBox("Can't Extract actual waist circumferance, Re-Draw Cut Out", 16, "Zipper Body")
                    Return False
                End If
                ScanLine(sTmp, nThighFiguredCir, nWaistActualCir)
                EOSFound = True
            Next

            If (EOSFound = False) Then
                MsgBox("Can't find an EOS for End of Zipper at the PROXIMAL EOS!", 16, "Zipper Body")
                Return False
            End If

            'Establish allowance for zippers
            nElasticProximal = 0.75
            If (String.Compare(Mid(sElasticProximal, 1, 1), "N") = 0) Then
                nElasticProximal = 0
            End If
            If (String.Compare(Mid(sElasticProximal, 1, 3), "3/8") = 0) Then
                nElasticProximal = 0.375
            End If
            If (String.Compare(Mid(sElasticProximal, 1, 1), "1") = 0) Then
                nElasticProximal = 1.5
            End If
            nElastic = nElasticProximal

            ARMDIA1.PR_SetLayer("Notes")
            'Draw Zipper
            'Establish Zipper Y
            Dim nInchToCM As Double = 1
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                nInchToCM = 2.54
            End If
            nSeam = 0.1875 * nInchToCM
            xyHorizEnd.y = xyOtemplate.y + nSeam
            xyHorizStart.y = xyHorizEnd.y

            If (EOStoSelectedPoint) Then
                xyHorizStart.X = xyHorizEnd.X
                While ((xyHorizEnd.X - xyHorizStart.X) <= 0)
                    Dim pPtRes As PromptPointResult
                    Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                    pPtOpts.Message = vbLf & "Select Start of Zipper "
                    pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                    If pPtRes.Status = PromptStatus.Cancel Then
                        MsgBox("Start not selected", 16, "Zipper Body")
                        Return False
                    End If
                    Dim ptStart As Point3d = pPtRes.Value
                    ARMDIA1.PR_MakeXY(xyHorizStart, ptStart.X, ptStart.Y)
                    If ((xyHorizEnd.X - xyHorizStart.X) < 0) Then
                        MsgBox("Select towards DISTAL end, Try again", 16, "Zipper Body")
                    End If
                End While
                xyHorizStart.y = xyHorizEnd.y
            End If

            If (nZipLength > 0) Then
                xyHorizStart.X = xyHorizEnd.X - ((nZipLength / 1.2) - nElastic)
            Else
                nDrawnLength = ARMDIA1.FN_CalcLength(xyHorizStart, xyHorizEnd)
                nZipLength = (nDrawnLength + nElastic) * 1.2
            End If

            If (nZipLength > 0) Then
                xyHorizStart.X = xyHorizEnd.X - ((nZipLength / 1.2) - nElastic)
            Else
                nDrawnLength = ARMDIA1.FN_CalcLength(xyHorizStart, xyHorizEnd)
                nZipLength = (nDrawnLength + nElastic) * 1.2
            End If
        End Using

        'Draw markers only
        ARMDIA1.PR_SetLayer("Notes")
        'Add Label And arrows
        PR_DrawClosedArrow(xyHorizStart, 0)
        PR_DrawClosedArrow(xyHorizEnd, 180)
        ''sZipLength = Format("length", nZipLength)
        sZipLength = Trim(LGLEGDIA1.fnInchesToText(nZipLength)) + Chr(34)
        ''sZipOffset = Format("length", nWaistActualCir / 6)
        sZipOffset = Trim(LGLEGDIA1.fnInchesToText(nWaistActualCir / 6)) + Chr(34)
        If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
            sZipLength = fnDisplayToCM(nZipLength)
            sZipOffset = fnDisplayToCM(nWaistActualCir / 6)
        End If

        'Add Text
        'Place relative to xyHorizEnd
        '  SetData("TextHorzJust", 1) ;
        '  SetData("TextVertJust", 32) ;
        '  SetData("TextHeight", 0.125) ;

        Dim xyText As ARMDIA1.XY
        ARMDIA1.PR_MakeXY(xyText, xyHorizEnd.X - 3, xyHorizEnd.y + 1.25)
        If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
            ARMDIA1.PR_MakeXY(xyText, xyHorizEnd.X - (3 * 2.54), xyHorizEnd.y + (1.25 * 2.54))
        End If
        ''PR_DrawText("Place " + sZipLength + " ZIPPER\n" + sZipOffset + " FROM CENTRE FRONT\n" + "SEAM ON LEFT SIDE", xyText, 0.125, 0)
        Dim sText As String = "Place " + sZipLength + " ZIPPER" + Chr(10) + sZipOffset + " FROM CENTRE FRONT" + Chr(10) + "SEAM ON LEFT SIDE"
        PR_DrawMText(sText, xyText, 0.125, 0)
        'Reset And exit
        MsgBox("Zipper drawing Complete", 0, "Zipper Body")
        Return True
    End Function
    Private Sub ScanLine(ByRef sLine As String, ByRef nNo As Decimal, ByRef sScale As Decimal)
        nNo = FN_GetNumber(sLine, 1)
        sScale = FN_GetNumber(sLine, 2)
    End Sub
    Public Function PR_EditZipper() As Boolean
        Dim sEntityID As String = ""
        Dim nFound As Integer = 0
        Dim sTmp As String = ""
        Dim sSymbolName As String = ""

        Dim sOldText As String = ""
        Dim sNewText As String = ""
        Dim idTextAsSymbol As ObjectId

        '     Clear CURRENT user selection
        'UserSelection("Clear");
        'UserSelection("Update");

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim ptCurveColl As Point3dCollection = New Point3dCollection
        Dim ptEntOpts As PromptEntityOptions = New PromptEntityOptions("Select Zipper")
        ptEntOpts.AllowNone = True
        Dim ptEntRes As PromptEntityResult = ed.GetEntity(ptEntOpts)
        If ptEntRes.Status <> PromptStatus.OK Then
            MsgBox("User cancelled", 0, "Edit Zipper")
            Return False
        End If

        Using tr As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim idObj As ObjectId = ptEntRes.ObjectId
            Dim acEnt As Entity = tr.GetObject(idObj, OpenMode.ForRead)
            Dim rb As ResultBuffer = acEnt.GetXDataForApplication("ID")
            If Not IsNothing(rb) Then
                ' Get the values in the xdata
                For Each typeVal As TypedValue In rb
                    If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                        Continue For
                    End If
                    sEntityID = sEntityID & typeVal.Value
                Next
            End If

            If (sEntityID.Length <= 0) Then
                MsgBox("Selected entity is not part of a Zipper!", 16, "Edit Zipper")
                Return False
            End If

            nFound = 0
            '' Get the Data From X marker
            Dim filterZipper(1) As TypedValue
            filterZipper.SetValue(New TypedValue(DxfCode.BlockName, "TEXTASSYMBOL"), 0)
            filterZipper.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "ID"), 1)
            Dim selFilterZipper As SelectionFilter = New SelectionFilter(filterZipper)
            Dim selResultZipper As PromptSelectionResult = ed.SelectAll(selFilterZipper)
            If selResultZipper.Status <> PromptStatus.OK Then
                MsgBox("Can't find Zipper Details", 16, "Edit Zipper")
                Return False
            End If
            Dim selSetZipper As SelectionSet = selResultZipper.Value

            Dim xyInsertPoint As ARMDIA1.XY
            For Each idObject As ObjectId In selSetZipper.GetObjectIds()
                Dim blkXMarker As BlockReference = tr.GetObject(idObject, OpenMode.ForRead)
                Dim rbID As ResultBuffer = blkXMarker.GetXDataForApplication("ID")
                Dim strID As String = ""
                If Not IsNothing(rbID) Then
                    ' Get the values in the xdata
                    For Each typeVal As TypedValue In rbID
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        strID = strID & typeVal.Value
                    Next
                End If
                If strID.Contains(sEntityID) = False Then
                    Continue For
                End If
                Dim position As Point3d = blkXMarker.Position
                ARMDIA1.PR_MakeXY(xyInsertPoint, position.X, position.Y)
                Dim resbuf As ResultBuffer = blkXMarker.GetXDataForApplication("Zipper")
                sTmp = ""
                If Not IsNothing(resbuf) Then
                    ' Get the values in the xdata
                    For Each typeVal As TypedValue In resbuf
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        sTmp = sTmp & typeVal.Value
                    Next
                End If

                If (String.Compare(sTmp, "1") = 0) Then
                    sSymbolName = blkXMarker.Name
                    If (sSymbolName.ToUpper().Equals("TEXTASSYMBOL", StringComparison.InvariantCultureIgnoreCase)) Then
                        idTextAsSymbol = idObject
                        Dim resbufData As ResultBuffer = blkXMarker.GetXDataForApplication("Data")
                        If Not IsNothing(resbufData) Then
                            ' Get the values in the xdata
                            For Each typeVal As TypedValue In resbufData
                                If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                    Continue For
                                End If
                                sOldText = sOldText & typeVal.Value
                            Next
                        End If
                    End If
                End If
            Next
            If (idTextAsSymbol.IsValid And (idTextAsSymbol.IsNull = False)) Then

                'nButX = 65; nButY = 35;
                '       hDlg = Open("dialog", "Zipper Text Editor", "font Helv 8", 20, 20, 215, 60);

                'AddControl(hDlg, "pushbutton", nButX, nButY, 35, 14, "Cancel", "%cancel", "");
                'AddControl(hDlg, "pushbutton", nButX + 48, nButY, 35, 14, "OK", "%ok", "");
                'AddControl(hDlg, "ledit", 5, 12, 204, 14, sOldText, "string", "sNewText");

                '  	Ok = Display("dialog", hDlg, "%center");
                ' 	Close("dialog", hDlg);
                Dim oZipEdit As New ZIPEDIT_frm
                oZipEdit.g_sCaption = "Zipper Text Editor"
                oZipEdit.g_sSymbolText = sOldText
                oZipEdit.ShowDialog()
                If oZipEdit.g_bIsCancel = True Then
                    Return False
                End If
                sNewText = oZipEdit.g_sSymbolText
            End If
            If (sOldText.Equals(sNewText) <> True) Then

                Dim blkXMarkerTemp As BlockReference = tr.GetObject(idTextAsSymbol, OpenMode.ForWrite)
                For Each attributeID As ObjectId In blkXMarkerTemp.AttributeCollection
                    Dim attRefObj As DBObject = tr.GetObject(attributeID, OpenMode.ForWrite)
                    If TypeOf attRefObj Is AttributeReference Then
                        Dim acAttRef As AttributeReference = attRefObj
                        If acAttRef.Tag.ToUpper().Equals("Data", StringComparison.InvariantCultureIgnoreCase) Then
                            acAttRef.TextString = sNewText
                        End If
                    End If
                Next
                Dim acRegAppTbl As RegAppTable
                acRegAppTbl = tr.GetObject(acCurDb.RegAppTableId, OpenMode.ForRead)
                Dim acRegAppTblRec As RegAppTableRecord
                If acRegAppTbl.Has("Data") = False Then
                    acRegAppTblRec = New RegAppTableRecord
                    acRegAppTblRec.Name = "Data"
                    acRegAppTbl.UpgradeOpen()
                    acRegAppTbl.Add(acRegAppTblRec)
                    tr.AddNewlyCreatedDBObject(acRegAppTblRec, True)
                End If
                Using rbNewData As New ResultBuffer
                    rbNewData.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, "Data"))
                    rbNewData.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, sNewText))
                    blkXMarkerTemp.XData = rbNewData
                End Using
                '' Save the new object to the database
                tr.Commit()
            End If
        End Using
        Return True
    End Function
    Public Function PR_DrawWaistLateralBodyZipper() As Boolean
        Dim sProfileID As String = ""
        Dim sStyleID As String = ""
        Dim sLeg As String = ""
        Dim nStringLength As Integer = 0
        Dim sOtemplateID As String = ""
        Dim sCO_WaistBottID As String = ""
        Dim sAnkleID As String = ""
        Dim nMarkersFound As Integer = 0
        Dim nType As Integer = 0
        Dim nn As Integer = 0
        Dim sTmp As String = ""
        Dim sUnits As String = ""
        Dim sType As String = ""
        Dim xyOtemplate As ARMDIA1.XY
        Dim xyCO_WaistBott As ARMDIA1.XY
        Dim xyAnkle As ARMDIA1.XY
        Dim xySmallestCir As ARMDIA1.XY
        Dim xyTmp As ARMDIA1.XY
        Dim xyCurveStart As ARMDIA1.XY

        Dim EOStoSelectedPoint As Boolean = False
        Dim EOStoEOS As Boolean = False
        Dim EOStoLegTape As Boolean = False

        Dim sDlgElasticList As String = ""
        Dim sDlgLengthList As String = ""
        Dim sDlgDistalElasticList As String = ""

        Dim nMedial As Integer = 0
        Dim bIsLoop As Boolean = False
        Dim nMinZipLength As Integer = 0
        Dim sZipLength As String = ""
        Dim nZipLength As Integer = 0
        Dim EOSFound As Boolean = False
        Dim xyMin As ARMDIA1.XY
        Dim xyMax As ARMDIA1.XY
        Dim xyHorizEnd As ARMDIA1.XY
        Dim xyHorizStart As ARMDIA1.XY
        Dim nThighFiguredCir As Integer = 0
        Dim nWaistActualCir As Integer = 0

        Dim nElasticProximal As Double = 0.0
        Dim sElasticProximal As String = ""
        Dim nElastic As Integer = 0

        Dim sElasticDistal As String = ""
        Dim nElasticDistal As Integer = 0
        Dim nTemplate As Integer = 0
        Dim sTemplate As String = ""
        Dim nTapeSpacing As Double = 0.0
        Dim nZipOff As Double = 0.0
        Dim nDrawnLength As Double = 0
        Dim sZipperID As String = ""


        '//Find JOBST installed directory
        '//Set path to macros

        '     String    sPathJOBST;
        'sPathJOBST = GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;
        'SetData("PathMacro", sPathJOBST + "\\WAIST");

        ''// Reset 
        '    UserSelection("clear");
        'UserSelection("update") ;
        'Execute("menu", "SetStyle", Table("find", "style", "bylayer"));
        'Execute("menu", "SetColor", Table("find", "color", "bylayer"));

        '// Get profile, identify Leg And FileNo
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim ptCurveColl As Point3dCollection = New Point3dCollection
        Dim ptEntOpts As PromptEntityOptions = New PromptEntityOptions("Select a Leg Profile")
        ptEntOpts.AllowNone = True
        ptEntOpts.SetRejectMessage("Only Curve")
        ptEntOpts.AddAllowedClass(GetType(Spline), True)
        Dim ptEntRes As PromptEntityResult = ed.GetEntity(ptEntOpts)
        If ptEntRes.Status <> PromptStatus.OK Then
            MsgBox("User cancelled", 0, "Lateral Body Zipper")
            Return False
        End If

        Using tr As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim idObj As ObjectId = ptEntRes.ObjectId
            Dim acSpline As Spline = tr.GetObject(idObj, OpenMode.ForRead)

            ''Dim acSpline As Spline = TryCast(acEnt, Spline)
            Dim fitPtsCount As Integer = acSpline.NumFitPoints
            Dim i As Integer = 0
            While (i < fitPtsCount)
                ptCurveColl.Add(acSpline.GetFitPointAt(i))
                i = i + 1
            End While
            Dim rb As ResultBuffer = acSpline.GetXDataForApplication("ID")
            If Not IsNothing(rb) Then
                ' Get the values in the xdata
                For Each typeVal As TypedValue In rb
                    If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                        Continue For
                    End If
                    sProfileID = sProfileID & typeVal.Value
                Next
            End If
            Dim bIsNotWaistLeg As Boolean = False
            Dim rbInsertion As ResultBuffer = acSpline.GetXDataForApplication("Insertion")
            If Not IsNothing(rbInsertion) Then
                ' Get the values in the xdata
                Dim strInsertion As String = ""
                For Each typeVal As TypedValue In rb
                    If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                        Continue For
                    End If
                    strInsertion = strInsertion & typeVal.Value
                Next
                If strInsertion <> "" Then
                    bIsNotWaistLeg = True
                End If
            End If

            nStringLength = sProfileID.Length
            If sProfileID.Contains("LeftLegCurve") Then
                sStyleID = Mid(sProfileID, 1, nStringLength - 12)
                sLeg = "Left"
            ElseIf sProfileID.Contains("RightLegCurve") Then
                sStyleID = Mid(sProfileID, 1, nStringLength - 13)
                sLeg = "Right"
            Else
                MsgBox("A Leg Profile was not selected", 16, "Lateral Body Zipper")
                Return False
            End If

            sOtemplateID = sStyleID + sLeg + "Origin"
            sAnkleID = sStyleID + sLeg + "Ankle"
            ''sStyleID = Mid(sStyleID, 5, sStyleID.Length - 4)
            Dim sCO_WaistStyleID As String = sStyleID
            If bIsNotWaistLeg = True Then
                sCO_WaistStyleID = Mid(sStyleID, 4, sStyleID.Length - 3)
            End If
            sCO_WaistBottID = sCO_WaistStyleID + sLeg + "CO_WaistBott"

            nMarkersFound = 0
            Dim Panty As Boolean = True
            '' Get the Data From X marker
            Dim filterMarker(4) As TypedValue
            filterMarker.SetValue(New TypedValue(DxfCode.BlockName, "X Marker"), 0)
            filterMarker.SetValue(New TypedValue(DxfCode.Operator, "<or"), 1)
            filterMarker.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "OriginID"), 2)
            filterMarker.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "ID"), 3)
            filterMarker.SetValue(New TypedValue(DxfCode.Operator, "or>"), 4)
            Dim selFilterMarker As SelectionFilter = New SelectionFilter(filterMarker)
            Dim selResultMarker As PromptSelectionResult = ed.SelectAll(selFilterMarker)
            If selResultMarker.Status <> PromptStatus.OK Then
                MsgBox("Can't find Marker Details", 16, "Lateral Body Zipper")
                Return False
            End If
            Dim selSetCLMarker As SelectionSet = selResultMarker.Value
            Dim xyInsertionXmarker As ARMDIA1.XY

            For Each idObject As ObjectId In selSetCLMarker.GetObjectIds()
                Dim blkXMarker As BlockReference = tr.GetObject(idObject, OpenMode.ForRead)
                Dim position As Point3d = blkXMarker.Position
                ARMDIA1.PR_MakeXY(xyInsertionXmarker, position.X, position.Y)
                If xyInsertionXmarker.X = 0 And xyInsertionXmarker.y = 0 Then
                    Continue For
                End If
                Dim resbuf As ResultBuffer = blkXMarker.GetXDataForApplication("OriginID")
                sTmp = ""
                If Not IsNothing(resbuf) Then
                    ' Get the values in the xdata
                    For Each typeVal As TypedValue In resbuf
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        sTmp = sTmp & typeVal.Value
                    Next
                End If

                If (String.Compare(sTmp, sOtemplateID) = 0) Then
                    nMarkersFound = nMarkersFound + 1
                    xyOtemplate = xyInsertionXmarker
                    Dim rbPressure As ResultBuffer = blkXMarker.GetXDataForApplication("Pressure")
                    If Not IsNothing(rbPressure) Then
                        ' Get the values in the xdata
                        For Each typeVal As TypedValue In rbPressure
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            sTemplate = sTemplate & typeVal.Value
                        Next
                    End If

                    Dim rbUnits As ResultBuffer = blkXMarker.GetXDataForApplication("units")
                    If Not IsNothing(rbUnits) Then
                        ' Get the values in the xdata
                        For Each typeVal As TypedValue In rbUnits
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            sUnits = sUnits & typeVal.Value
                        Next
                    End If
                    Dim rbData As ResultBuffer = blkXMarker.GetXDataForApplication("Data")
                    If Not IsNothing(rbData) Then
                        ' Get the values in the xdata
                        For Each typeVal As TypedValue In rbData
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            sTmp = sTmp & typeVal.Value
                        Next
                    End If

                    If (sTmp.Equals("")) Then
                        MsgBox("Can't get data from Origin Marker!", 16, "Lateral Body Zipper")
                        Return False
                    End If
                    ScanLine(sTmp, nType, nn)
                    sType = sTmp
                End If
                Dim resbufID As ResultBuffer = blkXMarker.GetXDataForApplication("ID")
                sTmp = ""
                If Not IsNothing(resbufID) Then
                    ' Get the values in the xdata
                    For Each typeVal As TypedValue In resbufID
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        sTmp = sTmp & typeVal.Value
                    Next
                End If
                If (String.Compare(sTmp, sAnkleID) = 0) Then
                    Panty = False
                    xyAnkle = xyInsertionXmarker
                End If
                If (String.Compare(sTmp, sCO_WaistBottID) = 0) Then
                    xyCO_WaistBott = xyInsertionXmarker
                End If
            Next

            'Check If that the markers have been found, otherwise exit
            If (String.Compare("CHAP", sType) = 0) Then
                MsgBox("Can't get data from Origin Marker!", 16, "Lateral Body Zipper")
                Return False
            End If

            If (nMarkersFound < 1) Then
                MsgBox("Missing marker for selected Waist Height, data not found!", 16, "Lateral Body Zipper")
                Return False
            End If

            If (nMarkersFound > 1) Then
                MsgBox("Two or more drawings of the same style exist!" & Chr(10) & "Delete the extra drawing/s and try again.", 16, "Lateral Body Zipper")
                Return False
            End If

            If (Panty) Then
                xySmallestCir.y = 10000   ''// Impossibly large value used as an initial value
                nn = 0
                While (nn < ptCurveColl.Count)
                    ARMDIA1.PR_MakeXY(xyTmp, ptCurveColl(nn).X, ptCurveColl(nn).Y)
                    If nn = 0 Then
                        xyCurveStart = xyTmp
                    End If
                    If (xyTmp.y < xySmallestCir.y) Then
                        xySmallestCir.y = xyTmp.y
                    Else
                        Exit While
                    End If
                    nn = nn + 1
                End While
            Else
                xySmallestCir.y = xyAnkle.y
            End If

            ' // Create Dialog
            '// Get Zipper style
            sUnits = "Inches"
            EOStoSelectedPoint = False
            EOStoEOS = False
            EOStoLegTape = False
            '         sDlgElasticList = "3/8\" Elastic\n3/4\" Elastic\n1½\" Elastic\nNo Elastic"
            'sDlgProximalElasticList =  "3/4\" Elastic\n" + sDlgElasticList
            'sDlgDistalElasticList = "No Elastic\n" + sDlgElasticList

            If (Panty) Then
                sDlgLengthList = "Leg Tape\nEOS\nSelected Point"
            Else
                sDlgLengthList = "Given a length\nSelected Point"
            End If

            nMedial = 0
            bIsLoop = True
            nMinZipLength = 5
            While (bIsLoop)
                '               nButX = 65; nButY = 45;
                '           hDlg = Open("dialog", sLeg + " Lateral Zipper (Waist Height)", "font Helv 8", 20, 20, 210, 65);

                'AddControl(hDlg, "pushbutton", nButX, nButY, 35, 14, "Cancel", "%cancel", "");
                'AddControl(hDlg, "pushbutton", nButX + 48, nButY, 35, 14, "OK", "%ok", "");

                'AddControl(hDlg, "ltext", 5, 12, 40, 14, "EOS to", "string", "");
                'AddControl(hDlg, "combobox", 35, 10, 70, 40, sDlgLengthList, "string", "sZipLength");
                'AddControl(hDlg, "ltext", 110, 12, 30, 14, "Proximal:", "string", "");
                '	AddControl(hDlg, "dropdownlist", 145, 10, 60, 70, sDlgProximalElasticList, "string", "sElasticProximal");
                'AddControl(hDlg, "ltext", 110, 30, 25, 14, "Distal :", "string", "");

                '	AddControl(hDlg, "dropdownlist", 145, 28, 60, 70, sDlgDistalElasticList, "string", "sElasticDistal");

                '  	Ok = Display("dialog", hDlg, "%center");
                ' 	Close("dialog", hDlg);

                'If (Ok == %cancel ) Exit (%ok, "User Cancel!") ;	
                Dim oZipperLaterFrm As New ZIPLAT_frm
                oZipperLaterFrm.g_sCaption = sLeg + " Lateral Zipper (Waist Height)"
                oZipperLaterFrm.g_bIsPanty = Panty
                oZipperLaterFrm.ShowDialog()
                If oZipperLaterFrm.g_bIsCancel = True Then
                    MsgBox("User Cancel!", 16, "Lateral Body Zipper")
                    Return False
                End If
                sZipLength = oZipperLaterFrm.g_sZipLength
                sElasticProximal = oZipperLaterFrm.g_sElasticProximal
                sElasticDistal = oZipperLaterFrm.g_sElasticDistal

                If (sZipLength.Equals("Leg Tape")) Then
                    EOStoLegTape = True
                End If
                If (sZipLength.Equals("EOS")) Then
                    EOStoEOS = True
                End If
                If (sZipLength.Equals("Selected Point")) Then
                    EOStoSelectedPoint = True
                End If
                If (EOStoLegTape = True Or EOStoEOS = True Or EOStoSelectedPoint = True) Then
                    bIsLoop = False
                Else
                    nZipLength = Val(sZipLength)
                    If (nZipLength = 0 And sZipLength.Length > 0) Then
                        MsgBox("Invalid given length!", 16, "Lateral Body Zipper")
                        bIsLoop = True
                    Else
                        bIsLoop = False
                    End If
                End If
            End While

            ' // Check that horizontal zipper line can intersect with the  EOS
            '// Get EOS at Proximal End
            EOSFound = False
            Dim filterLine(2) As TypedValue
            filterLine.SetValue(New TypedValue(DxfCode.Start, "Line"), 0)
            filterLine.SetValue(New TypedValue(DxfCode.LayerName, "Construct"), 1)
            ''filterLine.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, sStyleID + sLeg + "EOSLine"), 2)
            filterLine.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "EOSID"), 2)
            Dim selFilterLine As SelectionFilter = New SelectionFilter(filterLine)
            Dim selResultLine As PromptSelectionResult = ed.SelectAll(selFilterLine)
            If selResultLine.Status <> PromptStatus.OK Then
                MsgBox("Can't find an EOS for End of Zipper at the PROXIMAL EOS!", 16, "Lateral Body Zipper")
                Return False
            End If
            Dim selSetLine As SelectionSet = selResultLine.Value
            Dim extLine As Extents3d

            For Each idLine As ObjectId In selSetLine.GetObjectIds()
                Dim line As Line = tr.GetObject(idLine, OpenMode.ForRead)
                sTmp = ""
                Dim rbEOS As ResultBuffer = line.GetXDataForApplication("EOSID")
                If Not IsNothing(rbEOS) Then
                    ' Get the values in the xdata
                    For Each typeVal As TypedValue In rbEOS
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        sTmp = sTmp & typeVal.Value
                    Next
                End If
                If sTmp.Contains(sStyleID + sLeg + "EOSLine") = False Then
                    Continue For
                End If
                extLine = line.GeometricExtents
                Dim ptMin As Point3d = extLine.MinPoint
                Dim ptMax As Point3d = extLine.MaxPoint
                ARMDIA1.PR_MakeXY(xyMin, ptMin.X, ptMin.Y)
                ARMDIA1.PR_MakeXY(xyMax, ptMax.X, ptMax.Y)
                xyHorizEnd.X = xyMin.X
                sTmp = ""
                Dim rbLineData As ResultBuffer = line.GetXDataForApplication("Data")
                If Not IsNothing(rbLineData) Then
                    ' Get the values in the xdata
                    For Each typeVal As TypedValue In rbLineData
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        sTmp = sTmp & typeVal.Value
                    Next
                End If

                If (sTmp.Equals("")) Then
                    MsgBox("Can't Extract Figured Thigh Circumference or Actual Waist, Re-Draw Cut Out", 16, "Lateral Body Zipper")
                    Return False
                End If
                ScanLine(sTmp, nThighFiguredCir, nWaistActualCir)
                EOSFound = True
            Next

            If (EOSFound = False) Then
                MsgBox("Can't find an EOS for End of Zipper at the PROXIMAL EOS!", 16, "Lateral Body Zipper")
                Return False
            End If


            '//Establish allowance for zippers
            nElasticProximal = 0.75
            If (String.Compare(Mid(sElasticProximal, 1, 1), "N") = 0) Then
                nElasticProximal = 0
            End If
            If (String.Compare(Mid(sElasticProximal, 1, 3), "3/8") = 0) Then
                nElasticProximal = 0.375
            End If
            If (String.Compare(Mid(sElasticProximal, 1, 1), "1") = 0) Then
                nElasticProximal = 1.5
            End If

            If (EOStoSelectedPoint = True Or EOStoLegTape = True Or nZipLength > 0) Then ''//closed zipper at distal end
                nElasticDistal = 0
            Else
                nElasticDistal = 0.75
                If (String.Compare(Mid(sElasticDistal, 1, 1), "N") = 0) Then
                    nElasticDistal = 0
                End If
                If (String.Compare(Mid(sElasticDistal, 1, 3), "3/8") = 0) Then
                    nElasticDistal = 0.375
                End If
                If (String.Compare(Mid(sElasticDistal, 1, 1), "1") = 0) Then
                    nElasticDistal = 1.5
                End If
            End If
            nElastic = nElasticDistal + nElasticProximal

            '// Get template from leg box
            '// To establish Tape spacing 

            nTemplate = Val(Mid(sTemplate, 1, 2))
            If (nTemplate >= 30) Then
                nTapeSpacing = 1.25
            End If
            If (nTemplate = 13) Then
                nTapeSpacing = 1.31
            End If
            If (nTemplate = 9) Then
                nTapeSpacing = 1.37
            End If

            ' // Draw Zipper
            '// 
            '// Establish Zipper Y
            '// NB nThighFiguredCir Is given using the 1/2 scale.
            Dim nInchToCM As Double = 1
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                nInchToCM = 2.54
            End If
            nZipOff = 1.125 * nInchToCM
            If (Panty) Then
                xyHorizEnd.y = xyOtemplate.y + ((nThighFiguredCir * 2) / 6)
            Else
                xyHorizEnd.y = xyOtemplate.y + (nWaistActualCir / 6)
            End If

            If (xyHorizEnd.y > xySmallestCir.y - nZipOff) Then
                xyHorizEnd.y = xySmallestCir.y - nZipOff
                xyHorizStart.y = xyHorizEnd.y
            End If

            '// For Medial Zipper step back 3 tapes from groin height
            '// Establish X of start And end
            If (EOStoLegTape) Then
                xyHorizStart.X = xyCurveStart.X + nTapeSpacing
            End If
            If (EOStoEOS) Then
                xyHorizStart.X = xyCurveStart.X
            End If

            If (EOStoSelectedPoint) Then
                xyHorizStart.X = xyHorizEnd.X
                While ((xyHorizEnd.X - xyHorizStart.X) <= 0)
                    Dim pPtRes As PromptPointResult
                    Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                    pPtOpts.Message = vbLf & "Select Start of Zipper "
                    pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                    If pPtRes.Status = PromptStatus.Cancel Then
                        MsgBox("Start not selected", 16, "Lateral Body Zipper")
                        Return False
                    End If
                    Dim ptStart As Point3d = pPtRes.Value
                    ARMDIA1.PR_MakeXY(xyHorizStart, ptStart.X, ptStart.Y)
                    If ((xyHorizEnd.X - xyHorizStart.X) < 0) Then
                        MsgBox("Select towards DISTAL end, Try again", 16, "Lateral Body Zipper")
                    End If
                End While
                xyHorizStart.y = xyHorizEnd.y
            End If

            If (nZipLength > 0) Then
                xyHorizStart.X = xyHorizEnd.X - ((nZipLength / 1.2) - nElastic)
            Else
                nDrawnLength = ARMDIA1.FN_CalcLength(xyHorizStart, xyHorizEnd)
                nZipLength = (nDrawnLength + nElastic) * 1.2
            End If
        End Using

        '// Draw Zip
        ARMDIA1.PR_SetLayer("Notes")
        Dim sText As String = Trim(LGLEGDIA1.fnInchesToText(nZipLength)) + Chr(34)
        If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
            sText = fnDisplayToCM(nZipLength)
        End If
        Dim xyBlock As ARMDIA1.XY
        ARMDIA1.PR_MakeXY(xyBlock, xyHorizStart.X + (xyHorizEnd.X - xyHorizStart.X) / 2, xyHorizStart.y + 0.25)
        '// Create ID string
        sZipperID = sStyleID ''+ MakeString("scalar", UID("get", hEnt)) ;
        PR_DrawTextAsSymbol(xyBlock, sZipperID, sText + " LATERAL ZIPPER", 0)

        SetDBData("ID", sZipperID)
        SetDBData("Zipper", "1")
        SetDBData("Data", sText + " LATERAL ZIPPER")

        '// Draw zip line
        PR_DrawLine(xyHorizStart, xyHorizEnd)
        SetDBData("ID", sZipperID)
        SetDBData("Zipper", "1")

        '// Add label And arrows
        PR_DrawClosedArrow(xyHorizStart, 0)
        SetDBData("ID", sZipperID)
        SetDBData("Zipper", "1")
        If (xyHorizEnd.y > xyCO_WaistBott.y And xyCO_WaistBott.y <> 0 And xyCO_WaistBott.X <> 0) Then
            Dim xyArrow As ARMDIA1.XY
            ARMDIA1.PR_MakeXY(xyArrow, xyHorizEnd.X, xyCO_WaistBott.y - 0.025)
            PR_DrawClosedArrow(xyArrow, 225)
        Else
            PR_DrawClosedArrow(xyHorizEnd, 180)
        End If
        SetDBData("ID", sZipperID)
        SetDBData("Zipper", "1")

        ''// Reset And exit
        ''Execute("menu", "SetLayer", Table("find", "layer", "1")) ;
        MsgBox("Zipper drawing Complete", 0, "Lateral Body Zipper")
        Return True
    End Function
    Function SetDBData(ByVal sField As String, ByVal sValue As String) As Boolean
        If (idLastCreated.IsValid = False) Then
            Return False
        End If

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        '' Start a transaction
        Using tr As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acEnt As DBObject = tr.GetObject(idLastCreated, OpenMode.ForWrite)
            Dim acRegAppTbl As RegAppTable
            acRegAppTbl = tr.GetObject(acCurDb.RegAppTableId, OpenMode.ForRead)
            Dim acRegAppTblRec As RegAppTableRecord
            If acRegAppTbl.Has(sField) = False Then
                acRegAppTblRec = New RegAppTableRecord
                acRegAppTblRec.Name = sField
                acRegAppTbl.UpgradeOpen()
                acRegAppTbl.Add(acRegAppTblRec)
                tr.AddNewlyCreatedDBObject(acRegAppTblRec, True)
            End If
            Using rbNewData As New ResultBuffer
                rbNewData.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, sField))
                rbNewData.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, sValue))
                acEnt.XData = rbNewData
            End Using
            '' Save the new object to the database
            tr.Commit()
        End Using
    End Function
    Sub PR_DrawLine(ByRef xyStart As ARMDIA1.XY, ByRef xyFinish As ARMDIA1.XY)
        'To the DRAFIX macro file (given by the global fNum).
        'Write the syntax to draw a LINE between two points.
        'For this to work it assumes that the following DRAFIX variables
        'are defined
        '    XY      xyStart
        '    HANDLE  hEnt
        '
        'Note:-
        '    fNum, CC, QQ, NL are globals initialised by FN_Open
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

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
            Dim acLine As Line = New Line(New Point3d(xyStart.X, xyStart.y, 0),
                                    New Point3d(xyFinish.X, xyFinish.y, 0))
            'If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
            '    acLine.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.y, 0)))
            'End If

            '' Add the new object to the block table record and the transaction
            acBlkTblRec.AppendEntity(acLine)
            acTrans.AddNewlyCreatedDBObject(acLine, True)
            idLastCreated = acLine.ObjectId()

            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Private Sub PR_DrawTextAsSymbol(ByRef xyPoint As ARMDIA1.XY, ByRef sID As String, ByRef sData As String, ByRef nAngle As Double)
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim sName As String = "TEXTASSYMBOL"
        Dim strTag(2), strTextString(2) As String
        strTag(1) = "ID"
        strTag(2) = "Data"
        strTextString(1) = sID
        strTextString(2) = sData
        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim blkRecId As ObjectId = ObjectId.Null
            If Not acBlkTbl.Has(sName) Then
                Dim blkTblRecSymbol As BlockTableRecord = New BlockTableRecord()
                blkTblRecSymbol.Name = sName
                Dim acAttDef As New AttributeDefinition
                Dim ii As Double
                For ii = 1 To 2
                    acAttDef = New AttributeDefinition
                    acAttDef.Position = New Point3d(xyPoint.X, xyPoint.y, 0)
                    acAttDef.Prompt = strTag(ii)
                    acAttDef.Tag = strTag(ii)
                    acAttDef.TextString = strTextString(ii)
                    acAttDef.Height = 0.125
                    acAttDef.Justify = AttachmentPoint.BottomCenter
                    If ii = 2 Then
                        acAttDef.Invisible = False
                    Else
                        acAttDef.Invisible = True
                    End If
                    blkTblRecSymbol.AppendEntity(acAttDef)
                Next
                acBlkTbl.UpgradeOpen()
                acBlkTbl.Add(blkTblRecSymbol)
                acTrans.AddNewlyCreatedDBObject(blkTblRecSymbol, True)
                blkRecId = blkTblRecSymbol.Id
            Else
                blkRecId = acBlkTbl(sName)
            End If
            ' Insert the block into the current space
            If blkRecId <> ObjectId.Null Then
                'Create new block reference 
                Dim blkRef As BlockReference = New BlockReference(New Point3d(xyPoint.X, xyPoint.y, 0), blkRecId)
                blkRef.Rotation = nAngle * (ARMDIA1.PI / 180)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    blkRef.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyPoint.X, xyPoint.y, 0)))
                End If
                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)
                acBlkTblRec.AppendEntity(blkRef)
                acTrans.AddNewlyCreatedDBObject(blkRef, True)
                idLastCreated = blkRef.ObjectId
                '' Open the Block table record Model space for write
                Dim blkTblRec As BlockTableRecord
                blkTblRec = acTrans.GetObject(blkRecId, OpenMode.ForRead)
                For Each objID As ObjectId In blkTblRec
                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                    If TypeOf dbObj Is AttributeDefinition Then
                        Dim acAttDef As AttributeDefinition = dbObj
                        If Not acAttDef.Constant Then
                            Dim acAttRef As New AttributeReference
                            acAttRef.SetAttributeFromBlock(acAttDef, blkRef.BlockTransform)
                            If acAttRef.Tag.ToUpper().Equals("ID", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = sID
                            ElseIf acAttRef.Tag.ToUpper().Equals("Data", StringComparison.InvariantCultureIgnoreCase) Then
                                acAttRef.TextString = sData
                            End If
                            blkRef.AttributeCollection.AppendAttribute(acAttRef)
                            acTrans.AddNewlyCreatedDBObject(acAttRef, True)
                        End If
                    End If
                Next
            End If
            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_DrawMText(ByRef sText As Object, ByRef xyInsert As ARMDIA1.XY, ByRef nHeight As Double, ByRef nAngle As Double)
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                     OpenMode.ForRead)
            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                                        OpenMode.ForWrite)
            Dim mtx As New MText()
            mtx.Location = New Point3d(xyInsert.X, xyInsert.y, 0)
            mtx.SetDatabaseDefaults()
            mtx.TextStyleId = acCurDb.Textstyle
            ' current text size
            mtx.TextHeight = nHeight
            ' current textstyle
            mtx.Width = 0.0
            mtx.Rotation = nAngle * (ARMDIA1.PI / 180)
            mtx.Contents = sText
            mtx.Attachment = AttachmentPoint.BottomLeft
            mtx.SetAttachmentMovingLocation(mtx.Attachment)
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                mtx.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsert.X, xyInsert.y, 0)))
            End If
            acBlkTblRec.AppendEntity(mtx)
            acTrans.AddNewlyCreatedDBObject(mtx, True)
            '' Save the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub
    Function fnDisplayToCM(ByRef nInches As Double) As String
        Dim nCMVal, nDec As Double
        Dim iInt As Short
        nCMVal = nInches * 2.54
        nCMVal = System.Math.Abs(nCMVal)
        iInt = Int(nCMVal)
        nDec = nCMVal - iInt
        If nDec >= 0.5 Then
            fnDisplayToCM = Str(ARMDIA1.round(nCMVal))
            Exit Function
        End If
        If iInt <> 0 Then
            nDec = (nCMVal - iInt) * 10
            nDec = ARMDIA1.round(nDec) / 10
            nCMVal = iInt + nDec
            fnDisplayToCM = Str(nCMVal)
        Else
            fnDisplayToCM = Str(nCMVal)
        End If
    End Function
    Public Sub PR_DrawAnkleZipper()
        Dim xyHeel, xyOtemplate, xyAnkle, xySmallestCir, xyConstruct, xyInt As ARMDIA1.XY
        Dim SmallHeel As Boolean = False
        Dim nType, nn, nX, nY, nZipLength As Double
        Dim sZipLength, sElastic As String
        Dim nHeelRad, nHeelOff, nZipOff, nEOSRad, aStart, aDelta, nElastic As Double
        Dim bIsDrawCurve As Boolean = False
        Dim xyCurveStart, xyHorizStart, xyHorizEnd, xyEOSCen As ARMDIA1.XY
        Dim nMedial As Double = 0

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim ptCurveColl As Point3dCollection = New Point3dCollection
        Dim ptEntOpts As PromptEntityOptions = New PromptEntityOptions(Chr(10) & "Select a Leg Profile")
        ptEntOpts.AllowNone = True
        ptEntOpts.SetRejectMessage(Chr(10) & "Only Curve")
        ptEntOpts.AddAllowedClass(GetType(Spline), True)
        Dim ptEntRes As PromptEntityResult = ed.GetEntity(ptEntOpts)
        If ptEntRes.Status <> PromptStatus.OK Then
            MsgBox("User cancelled", 0, "Ankle Zipper")
            Exit Sub
        End If

        Using tr As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim idSpline As ObjectId = ptEntRes.ObjectId
            Dim acEnt As Spline = tr.GetObject(idSpline, OpenMode.ForRead)
            Dim rb As ResultBuffer = acEnt.GetXDataForApplication("ID")
            Dim strProfile As String = ""
            If Not IsNothing(rb) Then
                ' Get the values in the xdata
                For Each typeVal As TypedValue In rb
                    If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                        Continue For
                    End If
                    strProfile = strProfile & typeVal.Value
                Next
            End If
            strProfile = Trim(strProfile)
            Dim sStyleID, sLeg As String
            If strProfile.Contains("LeftLegCurve") Then
                ''nLen = nLen + 11
                sStyleID = Mid(strProfile, 1, strProfile.Length - 12)
                sLeg = "Left"
            ElseIf strProfile.Contains("RightLegCurve") Then
                ''nLen = nLen + 12
                sStyleID = Mid(strProfile, 1, strProfile.Length - 13)
                sLeg = "Right"
            Else
                MsgBox("A Leg Profile was not selected" & Chr(10), 16, "Ankle Zipper")
                Exit Sub
            End If
            ''// Get Marker data
            Dim sHeelID As String = sStyleID + sLeg + "Heel"
            Dim sOriginXdata As String = sStyleID + sLeg + "Origin"
            Dim sAnkleID As String = sStyleID + sLeg + "Ankle"
            Dim filterType(5) As TypedValue
            filterType.SetValue(New TypedValue(DxfCode.BlockName, "X Marker"), 0)
            filterType.SetValue(New TypedValue(DxfCode.Operator, "<or"), 1)
            filterType.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "OriginID"), 2)
            filterType.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "HeelID"), 3)
            filterType.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "ID"), 4)
            filterType.SetValue(New TypedValue(DxfCode.Operator, "or>"), 5)
            Dim selFilter As SelectionFilter = New SelectionFilter(filterType)
            Dim selResult As PromptSelectionResult = ed.SelectAll(selFilter)
            If selResult.Status <> PromptStatus.OK Then
                MsgBox("Can't find Marker Details", 16, "Ankle Zipper")
                Exit Sub
            End If
            Dim selectionSet As SelectionSet = selResult.Value
            Dim nMarkersFound As Double = 0
            For Each idObject As ObjectId In selectionSet.GetObjectIds()
                Dim blkXMarker As BlockReference = tr.GetObject(idObject, OpenMode.ForRead)
                Dim position As Point3d = blkXMarker.Position
                If position.X = 0 And position.Y = 0 Then
                    Continue For
                End If
                Dim resbuf As ResultBuffer = blkXMarker.GetXDataForApplication("OriginID")
                Dim strXMarkerData As String = ""
                If Not IsNothing(resbuf) Then
                    For Each typeVal As TypedValue In resbuf
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        strXMarkerData = strXMarkerData & typeVal.Value
                    Next
                End If
                If strXMarkerData.Equals(sOriginXdata) Then
                    nMarkersFound = nMarkersFound + 1
                    ARMDIA1.PR_MakeXY(xyOtemplate, position.X, position.Y)
                    Dim resbufData As ResultBuffer = blkXMarker.GetXDataForApplication("Data")
                    strXMarkerData = ""
                    If Not IsNothing(resbufData) Then
                        For Each typeVal As TypedValue In resbufData
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            strXMarkerData = strXMarkerData & typeVal.Value
                        Next
                    End If
                    If strXMarkerData = "" Then
                        MsgBox("Can't get data from Origin Marker!", 16, "Ankle Zipper")
                        Exit Sub
                    End If
                    ScanLine(strXMarkerData, nType, nn)
                End If
                Dim rbHeelID As ResultBuffer = blkXMarker.GetXDataForApplication("HeelID")
                Dim strHeelID As String = ""
                If Not IsNothing(rbHeelID) Then
                    For Each typeVal As TypedValue In rbHeelID
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        strHeelID = strHeelID & typeVal.Value
                    Next
                End If
                If strHeelID.Equals(sHeelID) Then
                    nMarkersFound = nMarkersFound + 1
                    ARMDIA1.PR_MakeXY(xyHeel, position.X, position.Y)
                    Dim rbHeelData As ResultBuffer = blkXMarker.GetXDataForApplication("Data")
                    strHeelID = ""
                    If Not IsNothing(rbHeelData) Then
                        For Each typeVal As TypedValue In rbHeelData
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            strHeelID = strHeelID & typeVal.Value
                        Next
                    End If
                    Dim nSHeel As Integer = Val(strHeelID)
                    If nSHeel = 1 Then
                        SmallHeel = True
                    End If
                End If
                Dim rbAnkleID As ResultBuffer = blkXMarker.GetXDataForApplication("ID")
                Dim strAnkleID As String = ""
                If Not IsNothing(rbAnkleID) Then
                    For Each typeVal As TypedValue In rbAnkleID
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        If typeVal.TypeCode = DxfCode.ExtendedDataInteger16 Then
                            SmallHeel = (typeVal.Value = 1)
                            Continue For
                        End If
                        strAnkleID = strAnkleID & typeVal.Value
                    Next
                End If
                If strAnkleID.Equals(sAnkleID) Then
                    nMarkersFound = nMarkersFound + 1
                    ARMDIA1.PR_MakeXY(xyAnkle, position.X, position.Y)
                    Dim rbAnkleData As ResultBuffer = blkXMarker.GetXDataForApplication("Data")
                    strAnkleID = ""
                    If Not IsNothing(rbAnkleData) Then
                        For Each typeVal As TypedValue In rbAnkleData
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            strAnkleID = strAnkleID & typeVal.Value
                        Next
                        If strAnkleID = "" Then
                            MsgBox("Can't get data from Ankle Marker!", 16, "Ankle Zipper")
                            Exit Sub
                        End If
                        ScanLine(strAnkleID, nX, nY)
                    End If
                End If
            Next
            ''// Check if that the markers have been found, otherwise exit
            If (nMarkersFound < 3) Then
                MsgBox("Missing markers for selected foot, data not found!", 16, "Ankle Zipper")
                Exit Sub
            End If
            If (nMarkersFound > 3) Then
                MsgBox("Two or more drawings of the same style exist!" & Chr(10) & "Delete the extra drawing/s and try again.", 16, "Ankle Zipper")
                Exit Sub
            End If
            ''// Get Smallest circumference (by requesting) a point
            ''// Do some checking that the given point Is sensible
            Dim bIsLoop As Boolean = True
            While (bIsLoop)
                Dim pPtRes As PromptPointResult
                Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                pPtOpts.Message = vbLf & "Select smallest leg circumference"
                pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                If pPtRes.Status = PromptStatus.Cancel Then
                    MsgBox("User Cancelled", 16, "Ankle Zipper")
                    Exit Sub
                End If
                Dim ptStart As Point3d = pPtRes.Value
                ARMDIA1.PR_MakeXY(xySmallestCir, ptStart.X, ptStart.Y)
                If xySmallestCir.y > xyAnkle.y Or xySmallestCir.y < xyOtemplate.y Then
                    MsgBox("Selected point above Ankle or below edge of template. Try again", 16, "Zipper Body")
                Else
                    bIsLoop = False
                End If
            End While

            ''// Create Dialog
            ''// Get Zipper style
            ''//
            Dim sUnits As String = "Inches"
            Dim bIsEOS As Boolean = False
            bIsLoop = True
            Dim bIsSelect As Boolean = False
            '         sDlgElasticList = "3/8\" Elastic\n3/4\" Elastic\n1½\" Elastic\nNo Elastic";
            'if (nType == 2) sDlgElasticList =  "No Elastic\n" + sDlgElasticList;
            '	else sDlgElasticList =  "3/4\" Elastic\n" + sDlgElasticList;
            'sDlgLengthList = "EOS\nSelected Point\n";
            While (bIsLoop)
                '               nButX = 60; nButY = 45;
                '           hDlg = Open("dialog", sLeg + " Ankle Zipper (Lower Extremity)", "font Helv 8", 20, 20, 215, 70);

                'AddControl(hDlg, "pushbutton", nButX, nButY, 35, 14, "Cancel", "%cancel", "");
                'AddControl(hDlg, "pushbutton", nButX + 48, nButY, 35, 14, "OK", "%ok", "");

                'AddControl(hDlg, "ltext", 5, 12, 30, 14, "Ankle to", "string", "");
                'AddControl(hDlg, "combobox", 40, 10, 65, 30, sDlgLengthList, "string", "sZipLength");
                'AddControl(hDlg, "ltext", 113, 12, 33, 14, "Proximal:", "string", "");
                '	AddControl(hDlg, "dropdownlist", 145, 10, 60, 60, sDlgElasticList, "string", "sElastic");

                ' 	AddControl(hDlg, "checkbox", nButX + 12, nButY - 18, 60, 15, "Medial Zipper", "number", "nMedial");

                '  	Ok = Display("dialog", hDlg, "%center");
                ' 	Close("dialog", hDlg);

                'if (Ok == %cancel ) Exit (%ok, "User Cancel!") ;
                Dim oZipAnkFrm As New ZIPANK_frm
                oZipAnkFrm.g_sCaption = sLeg + " Ankle Zipper (Lower Extremity)"
                oZipAnkFrm.ShowDialog()
                If oZipAnkFrm.g_bIsCancel = True Then
                    MsgBox("User Cancel!", 16, "Anklet Zipper")
                    Exit Sub
                End If
                sZipLength = oZipAnkFrm.g_sZipLength
                sElastic = oZipAnkFrm.g_sElastic
                If oZipAnkFrm.g_bIsMedialZipper = True Then
                    nMedial = 1
                End If

                If sZipLength.Equals("EOS") Or sZipLength.Equals("Selected Point") Then
                    bIsLoop = False
                    If sZipLength.Equals("EOS") Then
                        bIsEOS = True
                    Else
                        bIsSelect = True
                    End If
                Else
                    nZipLength = Val(sZipLength)
                    If (nZipLength = 0 And sZipLength.Length > 0) Then
                        bIsLoop = True
                    Else
                        bIsLoop = False
                        bIsEOS = False
                        bIsSelect = False
                    End If
                End If
            End While
            ''// Draw Chosen Zipper style
            ''// 
            If (SmallHeel) Then
                nHeelRad = 1.1760709
                nHeelOff = 0.25
            Else
                nHeelRad = 1.5053069
                nHeelOff = 0.5
            End If
            Dim nInchToCM As Double = 1
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                nInchToCM = 2.54
            End If
            nHeelRad = nHeelRad * nInchToCM
            nHeelOff = nHeelOff * nInchToCM
            nZipOff = 1.125 * nInchToCM
            xyConstruct.X = xyAnkle.X + nX  ''// nX And nY stored relative position Of the 
            xyConstruct.y = xyAnkle.y + nY  ''// heel construction circle center point 
            Dim xyStart, xyEnd As ARMDIA1.XY
            ARMDIA1.PR_MakeXY(xyStart, xyHeel.X + nHeelOff, xyConstruct.y)
            ARMDIA1.PR_MakeXY(xyEnd, xyHeel.X + nHeelOff, 0)
            Dim Ok As Short = FN_CirLinInt(xyStart, xyEnd, xyConstruct, nHeelRad + nZipOff, xyInt)
            If Ok Then
                ''// Draw polyline curve from intersection
                xyCurveStart = xyInt
                xyHorizStart.X = xySmallestCir.X
                xyHorizStart.y = xySmallestCir.y - nZipOff
                bIsDrawCurve = True
            Else
                ''// Otherwise degenerate to a straight line
                xyHorizStart.X = xyHeel.X + nHeelOff
                xyHorizStart.y = xyHeel.y - nZipOff
                bIsDrawCurve = False
            End If

            ''// Check that horizontal zipper line can intersect with the  EOS
            If bIsEOS Then
                Dim EOSFound As Boolean = False
                Dim sClosingLineID As String = sStyleID + sLeg + "ClosingLine"
                Dim filterTypeCL(4) As TypedValue
                filterTypeCL.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "ID"), 0)
                filterTypeCL.SetValue(New TypedValue(DxfCode.Operator, "<or"), 1)
                filterTypeCL.SetValue(New TypedValue(DxfCode.Start, "Line"), 2)
                filterTypeCL.SetValue(New TypedValue(DxfCode.Start, "Arc"), 3)
                filterTypeCL.SetValue(New TypedValue(DxfCode.Operator, "or>"), 4)
                Dim selFilterCL As SelectionFilter = New SelectionFilter(filterTypeCL)
                Dim selResultCL As PromptSelectionResult = ed.SelectAll(selFilterCL)
                If selResult.Status <> PromptStatus.OK Then
                    MsgBox("Can't find Marker Details", 16, "Ankle Zipper")
                    Exit Sub
                End If
                Dim selectionSetCL As SelectionSet = selResultCL.Value
                Dim extLine As Extents3d
                Dim xyMin, xyMax As ARMDIA1.XY
                For Each idObject As ObjectId In selectionSetCL.GetObjectIds()
                    Dim acEntity As Entity = tr.GetObject(idObject, OpenMode.ForRead)
                    extLine = acEntity.GeometricExtents
                    Dim ptMin As Point3d = extLine.MinPoint
                    Dim ptMax As Point3d = extLine.MaxPoint
                    ARMDIA1.PR_MakeXY(xyMin, ptMin.X, ptMin.Y)
                    ARMDIA1.PR_MakeXY(xyMax, ptMax.X, ptMax.Y)
                    If (xyHorizStart.y >= xyMin.y And xyHorizStart.y <= xyMax.y) Then
                        If (acEntity.GetType().Name.ToUpper.Equals("ARC")) Then
                            Dim acArc As Arc = TryCast(acEntity, Arc)
                            Dim ptCenter As Point3d = acArc.Center
                            ARMDIA1.PR_MakeXY(xyEOSCen, ptCenter.X, ptCenter.Y)
                            nEOSRad = acArc.Radius
                            aStart = acArc.StartAngle
                            Dim nEndAngle As Double = acArc.EndAngle
                            aDelta = nEndAngle - aStart
                            xyHorizEnd.X = xyEOSCen.X + System.Math.Sqrt(nEOSRad ^ 2 - (xyHorizStart.y - xyEOSCen.y) ^ 2)
                        End If
                        If (acEntity.GetType().Name.ToUpper.Equals("LINE")) Then
                            If xyMin.X = xyMax.X Then
                                xyHorizEnd.X = xyMin.X
                            Else
                                Dim acLine As Line = TryCast(acEntity, Line)
                                Dim ptStart As Point3d = acLine.StartPoint
                                Dim ptEnd As Point3d = acLine.EndPoint
                                Dim nTanTheta As Double = (ptEnd.Y - ptStart.Y) / (ptEnd.X - ptStart.X)
                                xyHorizEnd.X = xyMin.X + (xyHorizStart.y - xyMin.y) / nTanTheta
                            End If
                        End If
                        EOSFound = True
                        xyHorizEnd.y = xyHorizStart.y
                    End If
                Next
                If EOSFound = False Then
                    MsgBox("Can't find an EOS for zipper line to intersect!" & Chr(10) & "Select a point manually!", 16, "Ankle Zipper")
                    bIsSelect = True
                End If
                ''// Elastic allowance at EOS
                nElastic = 0.75
                If (String.Compare(Mid(sElastic, 1, 1), "N") = 0) Then
                    nElastic = 0
                End If
                If (String.Compare(Mid(sElastic, 1, 3), "3/8") = 0) Then
                    nElastic = 0.375
                End If
                If (String.Compare(Mid(sElastic, 1, 1), "1") = 0) Then
                    nElastic = 1.5
                End If
            Else
                ''// No elastic as this implies a closed zipper
                nElastic = 0
            End If

            ''// If select given then prompt user for end of zip
            ''//
            If bIsSelect Then
                Dim pEndPtRes As PromptPointResult
                Dim pEndPtOpts As PromptPointOptions = New PromptPointOptions("")
                pEndPtOpts.Message = vbLf & "Select End of Zipper"
                pEndPtRes = acDoc.Editor.GetPoint(pEndPtOpts)
                If pEndPtRes.Status = PromptStatus.Cancel Then
                    MsgBox("End not selected", 16, "Ankle Zipper")
                    Exit Sub
                End If
                Dim ptEndZip As Point3d = pEndPtRes.Value
                ARMDIA1.PR_MakeXY(xyHorizEnd, ptEndZip.X, ptEndZip.Y)
                xyHorizEnd.y = xyHorizStart.y
            End If
        End Using
        ''// Revise radius to include Zip offset
        ''//
        nHeelRad = nHeelRad + nZipOff
        ARMDIA1.PR_SetLayer("Notes")
        Dim nDrawnLength As Double
        If bIsDrawCurve Then
            ''StartPoly("fitted");
            ''AddVertex(xyCurveStart);
            ptCurveColl.Add(New Point3d(xyCurveStart.X, xyCurveStart.y, 0))
            Dim aAngle As Double = ARMDIA1.FN_CalcAngle(xyConstruct, xyHorizStart)
            Dim aPrevAngle As Double = ARMDIA1.FN_CalcAngle(xyConstruct, xyCurveStart)
            nDrawnLength = (aAngle - aPrevAngle) * ARMDIA1.PI * nHeelRad / 180
            Dim aAngleInc, ii As Double
            If aAngle > aPrevAngle Then
                aAngleInc = (aAngle - aPrevAngle) / 4
            Else
                aAngleInc = ((aAngle + 360) - aPrevAngle) / 4
            End If
            ii = 1
            Dim xyTmp, xyEnd As ARMDIA1.XY
            While (ii <= 3)
                ARMDIA1.PR_CalcPolar(xyConstruct, aPrevAngle + aAngleInc * ii, nHeelRad, xyTmp)
                If xyTmp.X < xyAnkle.X Then
                    ptCurveColl.Add(New Point3d(xyTmp.X, xyTmp.y, 0))
                    xyEnd = xyTmp
                End If
                ii = ii + 1
            End While
            ptCurveColl.Add(New Point3d(xyHorizStart.X, xyHorizStart.y, 0))
            ptCurveColl.Add(New Point3d(xyHorizStart.X + 0.25, xyHorizStart.y, 0))
            ptCurveColl.Add(New Point3d(xyHorizStart.X + 0.5, xyHorizStart.y, 0))

            ''// Requested zipper length
            If nZipLength > 0 Then
                xyHorizEnd.X = xyHorizStart.X + ((nZipLength / 1.2) - nDrawnLength)
                xyHorizEnd.y = xyHorizStart.y
            End If
            ptCurveColl.Add(New Point3d(xyHorizEnd.X, xyHorizEnd.y, 0))
            Dim nCurveLength = 0
            PR_DrawZipperSpline(ptCurveColl, nCurveLength)
            ''hEnt = UID("find", UID("getmax")) ;	
            SetDBData("ZipperLength", Str(nCurveLength))
            PR_DrawClosedArrow(xyCurveStart, aPrevAngle + 92)
            If nZipLength <= 0 Then
                ''----GetDBValue (hEnt, "ZipperLength", &sTmp, &nDrawnLength)
                nDrawnLength = nCurveLength
                If nType = 0 Then ''//Anklet only
                    ''SetData("TextHorzJust", 4) ;   	
                    nZipLength = nDrawnLength + nElastic
                Else
                    ''SetData("TextHorzJust", 2) ;   
                    nZipLength = (nDrawnLength + nElastic) * 1.2
                End If
            End If
        Else
            If nZipLength > 0 Then
                xyHorizEnd.X = xyHorizStart.X + ((nZipLength / 1.2))
                xyHorizEnd.y = xyHorizStart.y
            Else
                nDrawnLength = ARMDIA1.FN_CalcAngle(xyHorizStart, xyHorizEnd)
                If nType = 0 Then ''//Anklet only
                    ''SetData("TextHorzJust", 4) ;   	
                    nZipLength = nDrawnLength + nElastic
                Else
                    ''SetData("TextHorzJust", 2) ;   
                    nZipLength = (nDrawnLength + nElastic) * 1.2
                End If
            End If
            PR_DrawLine(xyHorizStart, xyHorizEnd)
            PR_DrawClosedArrow(xyHorizStart, 0)
        End If
        ''// Add label and arrows
        PR_DrawClosedArrow(xyHorizEnd, 180)
        Dim sZipperText As String = ""
        If nMedial = 1 Then
            sZipperText = " MEDIAL ZIPPER"
        Else
            sZipperText = " LATERAL ZIPPER"
        End If
        Dim sText As String = Trim(LGLEGDIA1.fnInchesToText(nZipLength)) + Chr(34) + sZipperText
        If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
            sText = fnDisplayToCM(nZipLength) + sZipperText
        End If
        ''SetData("TextVertJust", 32) ;
        ''SetData("TextHeight", 0.125) ;
        Dim xyText As ARMDIA1.XY
        ARMDIA1.PR_MakeXY(xyText, xyHorizStart.X + (xyHorizEnd.X - xyHorizStart.X) / 2, xyHorizStart.y + 0.25)
        PR_DrawZipperText(sText, xyText, 0.125, 0, 8)
        ''// Reset And exit
        ''Execute("menu", "SetLayer", Table("find", "layer", "1"))
        MsgBox("Zipper drawing Complete", 0, "Ankle Zipper")
    End Sub
    'To Draw Spline
    Sub PR_DrawZipperSpline(ByRef PointCollection As Point3dCollection, ByRef nCurveLength As Double)
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        '' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                                        OpenMode.ForWrite)
            '' Get a 3D vector from the point (0.5,0.5,0)
            Dim vecTan As Vector3d = New Point3d(0, 0, 0).GetAsVector
            '' Create a spline through (0, 0, 0), (5, 5, 0), and (10, 0, 0) with a
            '' start and end tangency of (0.5, 0.5, 0.0)
            Using acSpline As Spline = New Spline(PointCollection, vecTan, vecTan, 4, 0.0)
                'If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                '    acSpline.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyOtemplate.X, xyOtemplate.y, 0)))
                'End If
                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acSpline)
                acTrans.AddNewlyCreatedDBObject(acSpline, True)
                idLastCreated = acSpline.ObjectId
                nCurveLength = acSpline.GetDistanceAtParameter(acSpline.EndParam) - acSpline.GetDistanceAtParameter(acSpline.StartParam)
            End Using
            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Private Sub PR_DrawZipperText(ByRef sText As Object, ByRef xyInsert As ARMDIA1.XY, ByRef nHeight As Object, ByRef nAngle As Object, ByVal nTextmode As Double)
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
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)

            '' Create a single-line text object
            Using acText As DBText = New DBText()
                acText.Position = New Point3d(xyInsert.X, xyInsert.y, 0)
                acText.Height = nHeight
                acText.TextString = sText
                acText.Rotation = nAngle
                acText.WidthFactor = 0.6
                ''acText.HorizontalMode = nTextmode
                acText.Justify = nTextmode
                ''If acText.HorizontalMode <> TextHorizontalMode.TextLeft Then
                acText.AlignmentPoint = New Point3d(xyInsert.X, xyInsert.y, 0)
                ''End If
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acText.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsert.X, xyInsert.y, 0)))
                End If
                acBlkTblRec.AppendEntity(acText)
                acTrans.AddNewlyCreatedDBObject(acText, True)
            End Using

            '' Save the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub
    Public Sub PR_DrawDistalZipper()
        Dim sProfileID As String = ""
        Dim sStyleID As String = ""
        Dim sLeg As String = ""
        Dim nStringLength As Integer = 0
        Dim sOtemplateID As String = ""
        Dim sAnkleID As String = ""
        Dim nMarkersFound As Integer
        Dim sTmp As String = ""
        Dim xyOtemplate, xyAnkle, xySmallestCir, xyTmp, xyHorizStart, xyHorizEnd As ARMDIA1.XY
        Dim sUnits As String = ""
        Dim nType As Integer = 0
        Dim nn As Integer = 0
        Dim sType As String = ""
        Dim EOStoSelectedPoint, bIsLoop As Boolean
        Dim nMedial As Integer
        Dim sZipLength As String = ""
        Dim sElasticProximal As String = ""
        Dim sElasticDistal As String = ""
        Dim EOStoEOS As Boolean = False
        Dim nZipLength, nZipOff, nElasticProximal, nElasticDistal, nElastic, nDrawnLength As Double

        ''// Get profile, identify Leg and FileNo
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim ptCurveColl As Point3dCollection = New Point3dCollection
        Dim ptEntOpts As PromptEntityOptions = New PromptEntityOptions(Chr(20) & "Select a Leg Profile")
        ptEntOpts.AllowNone = True
        ptEntOpts.SetRejectMessage(Chr(20) & "Only Curve")
        ptEntOpts.AddAllowedClass(GetType(Spline), True)
        Dim ptEntRes As PromptEntityResult = ed.GetEntity(ptEntOpts)
        If ptEntRes.Status <> PromptStatus.OK Then
            MsgBox("User cancelled", 0, "Distal EOS Zipper")
            Exit Sub
        End If

        Using tr As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim idObj As ObjectId = ptEntRes.ObjectId
            Dim acSpline As Spline = tr.GetObject(idObj, OpenMode.ForRead)
            Dim fitPtsCount As Integer = acSpline.NumFitPoints
            Dim i As Integer = 0
            While (i < fitPtsCount)
                ptCurveColl.Add(acSpline.GetFitPointAt(i))
                i = i + 1
            End While
            Dim rb As ResultBuffer = acSpline.GetXDataForApplication("ID")
            If Not IsNothing(rb) Then
                ' Get the values in the xdata
                For Each typeVal As TypedValue In rb
                    If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                        Continue For
                    End If
                    sProfileID = sProfileID & typeVal.Value
                Next
            End If
            nStringLength = sProfileID.Length
            If sProfileID.Contains("LeftLegCurve") Then
                sStyleID = Mid(sProfileID, 1, nStringLength - 12)
                sLeg = "Left"
            ElseIf sProfileID.Contains("RightLegCurve") Then
                sStyleID = Mid(sProfileID, 1, nStringLength - 13)
                sLeg = "Right"
            Else
                MsgBox("A Leg Profile was not selected", 16, "Distal EOS Zipper")
                Exit Sub
            End If

            ''// Get Marker data
            sOtemplateID = sStyleID + sLeg + "Origin"
            sAnkleID = sStyleID + sLeg + "Ankle"
            nMarkersFound = 0
            Dim Panty As Boolean = True
            '' Get the Data From X marker
            Dim filterMarker(4) As TypedValue
            filterMarker.SetValue(New TypedValue(DxfCode.BlockName, "X Marker"), 0)
            filterMarker.SetValue(New TypedValue(DxfCode.Operator, "<or"), 1)
            filterMarker.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "OriginID"), 2)
            filterMarker.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "ID"), 3)
            filterMarker.SetValue(New TypedValue(DxfCode.Operator, "or>"), 4)
            Dim selFilterMarker As SelectionFilter = New SelectionFilter(filterMarker)
            Dim selResultMarker As PromptSelectionResult = ed.SelectAll(selFilterMarker)
            If selResultMarker.Status <> PromptStatus.OK Then
                MsgBox("Can't find Marker Details", 16, "Distal EOS Zipper")
                Exit Sub
            End If
            Dim selSetCLMarker As SelectionSet = selResultMarker.Value
            Dim xyInsertionXmarker As ARMDIA1.XY

            For Each idObject As ObjectId In selSetCLMarker.GetObjectIds()
                Dim blkXMarker As BlockReference = tr.GetObject(idObject, OpenMode.ForRead)
                Dim position As Point3d = blkXMarker.Position
                ARMDIA1.PR_MakeXY(xyInsertionXmarker, position.X, position.Y)
                If xyInsertionXmarker.X = 0 And xyInsertionXmarker.y = 0 Then
                    Continue For
                End If
                Dim resbuf As ResultBuffer = blkXMarker.GetXDataForApplication("OriginID")
                sTmp = ""
                If Not IsNothing(resbuf) Then
                    ' Get the values in the xdata
                    For Each typeVal As TypedValue In resbuf
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        sTmp = sTmp & typeVal.Value
                    Next
                End If
                If (String.Compare(sTmp, sOtemplateID) = 0) Then
                    nMarkersFound = nMarkersFound + 1
                    xyOtemplate = xyInsertionXmarker
                    Dim rbUnits As ResultBuffer = blkXMarker.GetXDataForApplication("units")
                    If Not IsNothing(rbUnits) Then
                        ' Get the values in the xdata
                        For Each typeVal As TypedValue In rbUnits
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            sUnits = sUnits & typeVal.Value
                        Next
                    End If
                    Dim rbData As ResultBuffer = blkXMarker.GetXDataForApplication("Data")
                    If Not IsNothing(rbData) Then
                        ' Get the values in the xdata
                        For Each typeVal As TypedValue In rbData
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            sTmp = sTmp & typeVal.Value
                        Next
                    End If

                    If (sTmp.Equals("")) Then
                        MsgBox("Can't get data from Origin Marker!", 16, "Distal EOS Zipper")
                        Exit Sub
                    End If
                    ScanLine(sTmp, nType, nn)
                    sType = sTmp
                End If
                Dim resbufID As ResultBuffer = blkXMarker.GetXDataForApplication("ID")
                sTmp = ""
                If Not IsNothing(resbufID) Then
                    ' Get the values in the xdata
                    For Each typeVal As TypedValue In resbufID
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        sTmp = sTmp & typeVal.Value
                    Next
                End If
                If (String.Compare(sTmp, sAnkleID) = 0) Then
                    Panty = False
                    xyAnkle = xyInsertionXmarker
                End If
            Next
            'Check If that the markers have been found, otherwise exit
            If (String.Compare("CHAP", sType) = 0) Then
                MsgBox("Can't use this zipper for a CHAP style!", 16, "Distal EOS Zipper")
                Exit Sub
            End If
            If (Panty = True And nMarkersFound < 1) Then
                MsgBox("Missing markers for selected style, data not found!", 16, "Distal EOS Zipper")
                Exit Sub
            End If
            If (Panty = True And nMarkersFound > 1) Then
                MsgBox("Two or more drawings of the same style exist!" & Chr(10) & "Delete the extra drawing/s and try again.", 16, "Distal EOS Zipper")
                Exit Sub
            End If
            If (Panty = False) Then
                MsgBox("Can only be used on Footless styles!", 16, "Distal EOS Zipper")
                Exit Sub
            End If

            ''// Get Smallest circumference 
            ''// Get Smallest circumference by looping through the vetex of the curve 
            xySmallestCir.y = 10000  ''// Impossibly large value used As an initial value  
            nn = 0
            While (nn < ptCurveColl.Count)
                ARMDIA1.PR_MakeXY(xyTmp, ptCurveColl(nn).X, ptCurveColl(nn).Y)
                If (xyTmp.y < xySmallestCir.y) Then
                    xySmallestCir.y = xyTmp.y
                Else
                    Exit While
                End If
                nn = nn + 1
            End While

            ''// Create Dialog
            ''//
            sUnits = "Inches"
            EOStoSelectedPoint = False

            bIsLoop = True
            Dim nMinZipLength As Double = 5

            'sDlgElasticList =  "3/8\" Elastic\n3/4\" Elastic\n1½\" Elastic\nNo Elastic";
            'sDlgProximalElasticList =  "No Elastic\n" + sDlgElasticList;
            'sDlgDistalElasticList =  "No Elastic\n" + sDlgElasticList;
            'sDlgLengthList =  "Give a length\nSelected Point";   
            nMedial = 0
            While (bIsLoop)
                'nButX = 70; nButY = 65;
                'hDlg = Open("dialog", sLeg + " Distal EOS Zipper (Waist Height)", "font Helv 8", 20, 20, 235, 95);

                'AddControl(hDlg, "pushbutton", nButX, nButY, 35, 14, "Cancel", "%cancel", "");
                'AddControl(hDlg, "pushbutton", nButX + 48, nButY, 35, 14, "OK", "%ok", "");
                'AddControl(hDlg, "ltext", 5, 12, 55, 14, "Distal EOS to", "string", "");
                'AddControl(hDlg, "combobox", 60, 10, 70, 40, sDlgLengthList, "string", "sZipLength");
                'AddControl(hDlg, "ltext", 135, 12, 25, 14, "Distal :", "string", "");
                'AddControl(hDlg, "dropdownlist", 170, 10, 60, 70, sDlgDistalElasticList, "string", "sElasticDistal");
                'AddControl(hDlg, "ltext", 135, 30, 30, 14, "Proximal:", "string", "");
                '	AddControl(hDlg, "dropdownlist", 170, 28, 60, 70, sDlgProximalElasticList, "string", "sElasticProximal");
                ' 	AddControl(hDlg, "checkbox", nButX + 12, nButY - 18, 65, 15, "Medial Zipper", "number", "nMedial");

                '  	Ok = Display("dialog", hDlg, "%center");
                ' 	Close("dialog", hDlg);

                'If (Ok == %cancel ) Exit (%ok, "User Cancel!") ;
                Dim oZipperDistalFrm As New ZIPDST_frm
                oZipperDistalFrm.g_sCaption = sLeg + " Distal EOS Zipper (Waist Height)"
                oZipperDistalFrm.ShowDialog()
                If oZipperDistalFrm.g_bIsCancel = True Then
                    MsgBox("User Cancel!", 16, "Distal EOS Zipper")
                    Exit Sub
                End If
                sZipLength = oZipperDistalFrm.g_sZipLength
                sElasticProximal = oZipperDistalFrm.g_sElasticProximal
                sElasticDistal = oZipperDistalFrm.g_sElasticDistal
                If oZipperDistalFrm.g_bIsMedialZipper = True Then
                    nMedial = 1
                End If

                If sZipLength.Equals("Selected Point") Then
                    EOStoSelectedPoint = True
                End If
                If EOStoEOS Or EOStoSelectedPoint Then
                    bIsLoop = False
                Else
                    nZipLength = Val(sZipLength)
                    If (nZipLength = 0 And sZipLength.Length > 0) Then
                        MsgBox("Invalid given length!", 16, "Distal EOS Zipper")
                        bIsLoop = True
                    Else
                        bIsLoop = False
                    End If
                End If
            End While

            ''// Start X and Y
            ''//
            Dim nInchToCM As Double = 1
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                nInchToCM = 2.54
            End If
            nZipOff = 1.125 * nInchToCM
            ARMDIA1.PR_MakeXY(xyHorizStart, ptCurveColl(0).X, ptCurveColl(0).Y)
            xyHorizStart.y = xySmallestCir.y - nZipOff
            xyHorizEnd.y = xyHorizStart.y
            ''// Establish allowance for zippers
            ''//
            nElasticProximal = 0.75
            If (String.Compare(Mid(sElasticProximal, 1, 1), "N") = 0) Then
                nElasticProximal = 0
            End If
            If (String.Compare(Mid(sElasticProximal, 1, 3), "3/8") = 0) Then
                nElasticProximal = 0.375
            End If
            If (String.Compare(Mid(sElasticProximal, 1, 1), "1") = 0) Then
                nElasticProximal = 1.5
            End If

            If (EOStoSelectedPoint = True Or nZipLength > 0) Then ''//closed zipper at distal end
                nElasticDistal = 0
            Else
                nElasticDistal = 0.75
                If (String.Compare(Mid(sElasticDistal, 1, 1), "N") = 0) Then
                    nElasticDistal = 0
                End If
                If (String.Compare(Mid(sElasticDistal, 1, 3), "3/8") = 0) Then
                    nElasticDistal = 0.375
                End If
                If (String.Compare(Mid(sElasticDistal, 1, 1), "1") = 0) Then
                    nElasticDistal = 1.5
                End If
            End If
            nElastic = nElasticDistal + nElasticProximal

            If EOStoSelectedPoint = True Then
                xyHorizEnd.X = xyHorizStart.X
                While ((xyHorizEnd.X - xyHorizStart.X) <= 0)
                    Dim pPtRes As PromptPointResult
                    Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                    pPtOpts.Message = vbLf & "Select End of Zipper"
                    pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                    If pPtRes.Status = PromptStatus.Cancel Then
                        MsgBox("End not selected", 16, "Distal EOS Zipper")
                        Exit Sub
                    End If
                    Dim ptStart As Point3d = pPtRes.Value
                    ARMDIA1.PR_MakeXY(xyHorizEnd, ptStart.X, ptStart.Y)
                    If ((xyHorizEnd.X - xyHorizStart.X) < 0) Then
                        MsgBox("Select towards PROXIMAL end, Try again", 16, "Distal EOS Zipper")
                    End If
                End While
                xyHorizEnd.y = xyHorizStart.y
            End If
            If (nZipLength > 0) Then
                xyHorizEnd.X = xyHorizStart.X + ((nZipLength / 1.2) - nElastic)
            Else
                nDrawnLength = ARMDIA1.FN_CalcLength(xyHorizStart, xyHorizEnd)
                nZipLength = (nDrawnLength + nElastic) * 1.2
            End If
        End Using
        ''// Draw on layer notes
        ARMDIA1.PR_SetLayer("Notes")
        ''// Add zipper 
        Dim sZipperText As String = ""
        If (nMedial = 1) Then
            sZipperText = " MEDIAL ZIPPER"
        Else
            sZipperText = " LATERAL ZIPPER"
        End If
        sTmp = Trim(LGLEGDIA1.fnInchesToText(nZipLength)) + Chr(34) + sZipperText
        If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
            sTmp = fnDisplayToCM(nZipLength) + sZipperText
        End If
        ''// Draw Zip text 
        Dim xyBlock As ARMDIA1.XY
        ARMDIA1.PR_MakeXY(xyBlock, xyHorizStart.X + (xyHorizEnd.X - xyHorizStart.X) / 2, xyHorizStart.y + 0.25)
        '// Create ID string
        Dim sZipperID As String = sStyleID ''+ MakeString("scalar", UID("get", hEnt)) ;
        PR_DrawTextAsSymbol(xyBlock, sZipperID, sTmp, 0)
        SetDBData("ID", sZipperID)
        SetDBData("Zipper", "1")
        SetDBData("Data", sTmp)

        ''// Draw zip line
        PR_DrawLine(xyHorizStart, xyHorizEnd)
        SetDBData("ID", sZipperID)
        SetDBData("Zipper", "1")

        '// Add label And arrows
        PR_DrawClosedArrow(xyHorizStart, 0)
        SetDBData("ID", sZipperID)
        SetDBData("Zipper", "1")

        PR_DrawClosedArrow(xyHorizEnd, 180)
        SetDBData("ID", sZipperID)
        SetDBData("Zipper", "1")

        ''// Reset And exit
        ''Execute("menu", "SetLayer", Table("find", "layer", "1")) ;
        MsgBox("Zipper drawing Complete", 0, "Distal EOS Zipper")
    End Sub
    Public Sub PR_DrawWaistAnkleZipper()
        Dim xyOtemplate, xyHeel, xyAnkle, xyCO_WaistBott, xySmallestCir, xyConstruct, xyInt As ARMDIA1.XY
        Dim sStyleID As String = ""
        Dim sUnits As String = ""
        Dim sTemplate As String = ""
        Dim sType As String = ""
        Dim SmallHeel As Boolean = False
        Dim nX, nY, nZipLength, nHeelRad, nHeelOff, nZipOff, nElastic, nTemplate, nFoldOffset As Double
        Dim nMedial As Integer
        Dim sZipLength As String = ""
        Dim sElastic As String = ""
        Dim xyCurveStart, xyHorizStart, xyHorizEnd, xyTmp, xyEnd As ARMDIA1.XY
        Dim EOS As Boolean = False
        Dim DrawCurve As Boolean = False
        Dim EOSFound As Boolean = False
        Dim aAngle, aPrevAngle, nDrawnLength, aAngleInc As Double

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim ptCurveColl As Point3dCollection = New Point3dCollection
        Dim ptEntOpts As PromptEntityOptions = New PromptEntityOptions(Chr(10) & "Select a Leg Profile")
        ptEntOpts.AllowNone = True
        ptEntOpts.SetRejectMessage(Chr(10) & "Only Curve")
        ptEntOpts.AddAllowedClass(GetType(Spline), True)
        Dim ptEntRes As PromptEntityResult = ed.GetEntity(ptEntOpts)
        If ptEntRes.Status <> PromptStatus.OK Then
            MsgBox("User cancelled", 0, "Ankle Zipper")
            Exit Sub
        End If

        Using tr As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim idSpline As ObjectId = ptEntRes.ObjectId
            Dim acEnt As Spline = tr.GetObject(idSpline, OpenMode.ForRead)
            Dim rb As ResultBuffer = acEnt.GetXDataForApplication("ID")
            Dim strProfile As String = ""
            If Not IsNothing(rb) Then
                ' Get the values in the xdata
                For Each typeVal As TypedValue In rb
                    If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                        Continue For
                    End If
                    strProfile = strProfile & typeVal.Value
                Next
            End If
            strProfile = Trim(strProfile)
            Dim sLeg As String
            If strProfile.Contains("LeftLegCurve") Then
                ''nLen = nLen + 11
                sStyleID = Mid(strProfile, 1, strProfile.Length - 12)
                sLeg = "Left"
            ElseIf strProfile.Contains("RightLegCurve") Then
                ''nLen = nLen + 12
                sStyleID = Mid(strProfile, 1, strProfile.Length - 13)
                sLeg = "Right"
            Else
                MsgBox("A Leg Profile was not selected" & Chr(10), 16, "Ankle Zipper")
                Exit Sub
            End If
            ''// Get Marker data
            Dim sHeelID As String = sStyleID + sLeg + "Heel"
            Dim sOtemplateID As String = sStyleID + sLeg + "Origin"
            Dim sAnkleID As String = sStyleID + sLeg + "Ankle"
            Dim sCO_WaistBottID As String = sStyleID + sLeg + "CO_WaistBott"
            Dim filterType(5) As TypedValue
            filterType.SetValue(New TypedValue(DxfCode.BlockName, "X Marker"), 0)
            filterType.SetValue(New TypedValue(DxfCode.Operator, "<or"), 1)
            filterType.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "OriginID"), 2)
            filterType.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "HeelID"), 3)
            filterType.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "ID"), 4)
            filterType.SetValue(New TypedValue(DxfCode.Operator, "or>"), 5)
            Dim selFilter As SelectionFilter = New SelectionFilter(filterType)
            Dim selResult As PromptSelectionResult = ed.SelectAll(selFilter)
            If selResult.Status <> PromptStatus.OK Then
                MsgBox("Can't find Marker Details", 16, "Ankle Zipper")
                Exit Sub
            End If
            Dim selectionSet As SelectionSet = selResult.Value
            Dim nMarkersFound As Double = 0
            For Each idObject As ObjectId In selectionSet.GetObjectIds()
                Dim blkXMarker As BlockReference = tr.GetObject(idObject, OpenMode.ForRead)
                Dim position As Point3d = blkXMarker.Position
                If position.X = 0 And position.Y = 0 Then
                    Continue For
                End If
                Dim resbuf As ResultBuffer = blkXMarker.GetXDataForApplication("OriginID")
                Dim strXMarkerData As String = ""
                If Not IsNothing(resbuf) Then
                    For Each typeVal As TypedValue In resbuf
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        strXMarkerData = strXMarkerData & typeVal.Value
                    Next
                End If
                If strXMarkerData.Equals(sOtemplateID) Then
                    nMarkersFound = nMarkersFound + 1
                    ARMDIA1.PR_MakeXY(xyOtemplate, position.X, position.Y)
                    Dim resbufUnits As ResultBuffer = blkXMarker.GetXDataForApplication("units")
                    sUnits = ""
                    If Not IsNothing(resbufUnits) Then
                        For Each typeVal As TypedValue In resbufUnits
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            sUnits = sUnits & typeVal.Value
                        Next
                    End If
                    Dim resbufPressure As ResultBuffer = blkXMarker.GetXDataForApplication("Pressure")
                    sTemplate = ""
                    If Not IsNothing(resbufPressure) Then
                        For Each typeVal As TypedValue In resbufPressure
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            sTemplate = sTemplate & typeVal.Value
                        Next
                    End If
                    Dim resbufData As ResultBuffer = blkXMarker.GetXDataForApplication("Data")
                    sType = ""
                    If Not IsNothing(resbufData) Then
                        For Each typeVal As TypedValue In resbufData
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            sType = sType & typeVal.Value
                        Next
                    End If
                End If
                Dim rbHeelID As ResultBuffer = blkXMarker.GetXDataForApplication("HeelID")
                Dim strHeelID As String = ""
                If Not IsNothing(rbHeelID) Then
                    For Each typeVal As TypedValue In rbHeelID
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        strHeelID = strHeelID & typeVal.Value
                    Next
                End If
                If strHeelID.Equals(sHeelID) Then
                    nMarkersFound = nMarkersFound + 1
                    ARMDIA1.PR_MakeXY(xyHeel, position.X, position.Y)
                    Dim rbHeelData As ResultBuffer = blkXMarker.GetXDataForApplication("Data")
                    strHeelID = ""
                    If Not IsNothing(rbHeelData) Then
                        For Each typeVal As TypedValue In rbHeelData
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            strHeelID = strHeelID & typeVal.Value
                        Next
                    End If
                    Dim nSHeel As Integer = Val(strHeelID)
                    If nSHeel = 1 Then
                        SmallHeel = True
                    End If
                End If
                Dim rbAnkleID As ResultBuffer = blkXMarker.GetXDataForApplication("ID")
                Dim strAnkleID As String = ""
                If Not IsNothing(rbAnkleID) Then
                    For Each typeVal As TypedValue In rbAnkleID
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        strAnkleID = strAnkleID & typeVal.Value
                    Next
                End If
                If strAnkleID.Equals(sAnkleID) Then
                    nMarkersFound = nMarkersFound + 1
                    ARMDIA1.PR_MakeXY(xyAnkle, position.X, position.Y)
                    Dim rbAnkleData As ResultBuffer = blkXMarker.GetXDataForApplication("Data")
                    strAnkleID = ""
                    If Not IsNothing(rbAnkleData) Then
                        For Each typeVal As TypedValue In rbAnkleData
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            strAnkleID = strAnkleID & typeVal.Value
                        Next
                        If strAnkleID = "" Then
                            MsgBox("Can't get data from Ankle Marker!", 16, "Ankle Zipper")
                            Exit Sub
                        End If
                        ScanLine(strAnkleID, nX, nY)
                    End If
                End If
                If strAnkleID.Equals(sCO_WaistBottID) Then
                    ARMDIA1.PR_MakeXY(xyCO_WaistBott, position.X, position.Y)
                End If
            Next
            ''// Check if that the markers have been found, otherwise exit
            If sType.Contains("CHAP") Then
                MsgBox("Can't use this zipper for a CHAP style!", 16, "Ankle Zipper")
                Exit Sub
            End If
            If (nMarkersFound < 3) Then
                MsgBox("Missing markers for selected foot, data not found!", 16, "Ankle Zipper")
                Exit Sub
            End If
            If (nMarkersFound > 3) Then
                MsgBox("Two or more drawings of the same style exist!" & Chr(10) & "Delete the extra drawing/s and try again.", 16, "Ankle Zipper")
                Exit Sub
            End If

            ''// Get Smallest circumference (by requesting) a point
            ''// Do some checking that the given point Is sensible
            Dim bIsLoop As Boolean = True
            While (bIsLoop)
                Dim pPtRes As PromptPointResult
                Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                pPtOpts.Message = vbLf & "Select smallest leg circumference"
                pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                If pPtRes.Status = PromptStatus.Cancel Then
                    MsgBox("User Cancelled", 16, "Ankle Zipper")
                    Exit Sub
                End If
                Dim ptStart As Point3d = pPtRes.Value
                ARMDIA1.PR_MakeXY(xySmallestCir, ptStart.X, ptStart.Y)
                If xySmallestCir.y > xyAnkle.y Or xySmallestCir.y < xyOtemplate.y Then
                    MsgBox("Selected point above Ankle or below edge of template. Try again", 16, "Ankle Zipper")
                Else
                    bIsLoop = False
                End If
            End While
            ''// Create Dialog
            ''// Get Zipper style
            ''//
            sUnits = "Inches"
            EOS = False
            bIsLoop = True
            Dim bIsSelect As Boolean = False
            'sDlgElasticList = "3/8\" Elastic\n3/4\" Elastic\n1½\" Elastic\nNo Elastic";
            'sDlgElasticList =  "3/4\" Elastic\n" + sDlgElasticList;
            'sDlgLengthList = "EOS\nGive a Length\nSelected Point\n";
            nMedial = 0
            While (bIsLoop)
                'nButX = 60; nButY = 45;
                'hDlg = Open("dialog", sLeg + " Ankle Zipper (Waist Height)", "font Helv 8", 20, 20, 215, 70);

                'AddControl(hDlg, "pushbutton", nButX, nButY, 35, 14, "Cancel", "%cancel", "");
                'AddControl(hDlg, "pushbutton", nButX + 48, nButY, 35, 14, "OK", "%ok", "");

                'AddControl(hDlg, "ltext", 5, 12, 30, 14, "Ankle to", "string", "");
                'AddControl(hDlg, "combobox", 40, 10, 65, 45, sDlgLengthList, "string", "sZipLength");
                'AddControl(hDlg, "ltext", 113, 12, 33, 14, "Proximal:", "string", "");
                '	AddControl(hDlg, "dropdownlist", 145, 10, 60, 60, sDlgElasticList, "string", "sElastic");

                ' 	AddControl(hDlg, "ltext", 5, nButY - 15, 60, 15, "Template: " + sTemplate, "String", "nMedial");
                ' 	AddControl(hDlg, "checkbox", 120, nButY - 18, 60, 15, "Medial Zipper", "number", "nMedial");

                '  	Ok = Display("dialog", hDlg, "%center");
                ' 	Close("dialog", hDlg);

                'If (Ok == %cancel ) Exit (%ok, "User Cancel!") ;

                Dim oWHZipAnkFrm As New WHZIPANK_frm
                oWHZipAnkFrm.g_sCaption = sLeg + " Ankle Zipper (Waist Height)"
                oWHZipAnkFrm.g_sTemplate = ""
                If sTemplate <> "" Then
                    oWHZipAnkFrm.g_sTemplate = "Template: " + sTemplate
                End If
                oWHZipAnkFrm.ShowDialog()
                If oWHZipAnkFrm.g_bIsCancel = True Then
                    MsgBox("User Cancel!", 16, "Anklet Zipper")
                    Exit Sub
                End If
                sZipLength = oWHZipAnkFrm.g_sZipLength
                sElastic = oWHZipAnkFrm.g_sElastic
                If oWHZipAnkFrm.g_bIsMedialZipper = True Then
                    nMedial = 1
                End If

                If sZipLength.Equals("EOS") Or sZipLength.Equals("Selected Point") Then
                    bIsLoop = False
                    If (sZipLength.Equals("EOS")) Then
                        EOS = True
                    Else
                        bIsSelect = True
                    End If
                Else
                    nZipLength = Val(sZipLength)
                    If (nZipLength = 0 And sZipLength.Length > 0) Then
                        bIsLoop = True
                    Else
                        bIsLoop = False
                        EOS = False
                        bIsSelect = False
                    End If
                End If
            End While
            ''// Draw Chosen Zipper style
            ''// 
            If (SmallHeel) Then
                nHeelRad = 1.1760709
                nHeelOff = 0.25
            Else
                nHeelRad = 1.5053069
                nHeelOff = 0.5
            End If
            Dim nInchToCM As Double = 1
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                nInchToCM = 2.54
            End If
            nZipOff = 1.125 * nInchToCM

            xyConstruct.x = xyAnkle.X + nX  ''// nX And nY stored relative position Of the 
            xyConstruct.y = xyAnkle.y + nY  ''// heel construction circle center point 
            Dim xyStart, xyEndPt As ARMDIA1.XY
            ARMDIA1.PR_MakeXY(xyStart, xyHeel.X + nHeelOff, xyConstruct.y)
            ARMDIA1.PR_MakeXY(xyEndPt, xyHeel.X + nHeelOff, 0)
            Dim Ok As Short = FN_CirLinInt(xyStart, xyEndPt, xyConstruct, nHeelRad + nZipOff, xyInt)
            If (Ok) Then
                ''// Draw polyline curve from intersection
                xyCurveStart = xyInt
                xyHorizStart.x = xySmallestCir.X
                xyHorizStart.y = xySmallestCir.y - nZipOff
                DrawCurve = True
            Else
                ''// Otherwise degenerate to a straight line
                xyHorizStart.x = xyHeel.X + nHeelOff
                xyHorizStart.y = xyHeel.y - nZipOff
                DrawCurve = False
            End If
            ''// Check that horizontal zipper line can intersect with the  EOS
            If (EOS) Then
                EOSFound = False
                ''Dim sFileNo As String = StringMiddle(sStyleID, 5, StringLength(sStyleID) - 4) ;
                Dim sTmp As String = ""
                Dim sXdataName As String = ""
                If (nMedial = 1) Then
                    sTmp = sStyleID + sLeg + "FoldLine"
                    sXdataName = "ID"
                Else
                    sTmp = sStyleID + sLeg + "EOSLine"
                    sXdataName = "EOSID"
                End If
                Dim filterTypeCL(2) As TypedValue
                filterTypeCL.SetValue(New TypedValue(DxfCode.Start, "Line"), 0)
                filterTypeCL.SetValue(New TypedValue(DxfCode.LayerName, "Construct"), 1)
                filterTypeCL.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, sXdataName), 2)
                Dim selFilterCL As SelectionFilter = New SelectionFilter(filterTypeCL)
                Dim selResultCL As PromptSelectionResult = ed.SelectAll(selFilterCL)
                If selResult.Status <> PromptStatus.OK Then
                    MsgBox("Can't find Marker Details", 16, "Ankle Zipper")
                    Exit Sub
                End If
                Dim selectionSetCL As SelectionSet = selResultCL.Value
                Dim extLine As Extents3d
                Dim xyMin, xyMax As ARMDIA1.XY
                For Each idObject As ObjectId In selectionSetCL.GetObjectIds()
                    Dim acLine As Line = tr.GetObject(idObject, OpenMode.ForRead)
                    extLine = acLine.GeometricExtents
                    Dim ptMin As Point3d = extLine.MinPoint
                    Dim ptMax As Point3d = extLine.MaxPoint
                    ARMDIA1.PR_MakeXY(xyMin, ptMin.X, ptMin.Y)
                    ARMDIA1.PR_MakeXY(xyMax, ptMax.X, ptMax.Y)
                    If (xyHorizStart.y >= xyMin.y And xyHorizStart.y <= xyMax.y) Then
                        If xyMin.X = xyMax.X Then
                            xyHorizEnd.X = xyMin.X
                        Else  ''// non vertical EOS
                            Dim ptStart As Point3d = acLine.StartPoint
                            Dim ptEnd As Point3d = acLine.EndPoint
                            Dim nTanTheta As Double = (ptEnd.Y - ptStart.Y) / (ptEnd.X - ptStart.X)
                            xyHorizEnd.X = xyMin.X + (xyHorizStart.y - xyMin.y) / nTanTheta
                        End If
                        EOSFound = True
                        xyHorizEnd.y = xyHorizStart.y
                    End If
                Next
                If (EOSFound = False) Then
                    MsgBox("Can't find an EOS for zipper line to intersect!" & Chr(10) & "Select a point manually!", 16, "Ankle Zipper")
                    bIsSelect = True
                End If
                If (nMedial = 1) Then ''// closed Then zipper
                    nElastic = 0
                Else
                    nElastic = 0.75
                    If (String.Compare(Mid(sElastic, 1, 1), "N") = 0) Then
                        nElastic = 0
                    End If
                    If (String.Compare(Mid(sElastic, 1, 3), "3/8") = 0) Then
                        nElastic = 0.375
                    End If
                    If (String.Compare(Mid(sElastic, 1, 1), "1") = 0) Then
                        nElastic = 1.5
                    End If
                End If
            Else
                ''// No elastic as this implies a closed zipper
                nElastic = 0
                xyHorizEnd.X = xyHorizStart.X + nZipLength      ''// Do this To give an approximate position
                ''// As we need xyHorizEnd to insert the text symbol 
            End If

            ''// For Medial Zipper step back 3 tape
            ''// Get template from leg box
            If (nMedial = 1) Then
                nTemplate = Val(Mid(sTemplate, 1, 2))
                If (nTemplate >= 30) Then nFoldOffset = 1.25 * 3
                If (nTemplate = 13) Then nFoldOffset = 1.31 * 3
                If (nTemplate = 9) Then nFoldOffset = 1.37 * 3
                xyHorizEnd.X = xyHorizEnd.X - nFoldOffset
            End If

            ''// If select given then prompt user for end of zip
            ''//
            If (bIsSelect) Then
                Dim pEndPtRes As PromptPointResult
                Dim pEndPtOpts As PromptPointOptions = New PromptPointOptions("")
                pEndPtOpts.Message = vbLf & "Select End of Zipper"
                pEndPtRes = acDoc.Editor.GetPoint(pEndPtOpts)
                If pEndPtRes.Status = PromptStatus.Cancel Then
                    MsgBox("End not selected", 16, "Ankle Zipper")
                    Exit Sub
                End If
                Dim ptEndZip As Point3d = pEndPtRes.Value
                ARMDIA1.PR_MakeXY(xyHorizEnd, ptEndZip.X, ptEndZip.Y)
                xyHorizEnd.y = xyHorizStart.y
            End If
        End Using
        '' // Revise radius to include Zip offset
        ''//
        nHeelRad = nHeelRad + nZipOff
        ARMDIA1.PR_SetLayer("Notes")
        Dim xyBlock As ARMDIA1.XY
        ARMDIA1.PR_MakeXY(xyBlock, xyHorizStart.X + (xyHorizEnd.X - xyHorizStart.X) / 2, xyHorizStart.y + 0.125)
        '// Create ID string
        Dim sZipperID As String = sStyleID ''+ MakeString("scalar", UID("get", hEnt)) ;
        'PR_DrawTextAsSymbol(xyBlock, sZipperID, "", 0)
        'Dim idTextAsSymbol As ObjectId = idLastCreated
        'SetDBData("ID", sZipperID)
        'SetDBData("Zipper", "1")
        'SetDBData("Data", "INVALID ZIPPER CALCULATION") ''//Do this just incase it fails
        If DrawCurve = True Then
            ''StartPoly("fitted");
            ''AddVertex(xyCurveStart);
            ptCurveColl.Add(New Point3d(xyCurveStart.X, xyCurveStart.y, 0))
            aAngle = ARMDIA1.FN_CalcAngle(xyConstruct, xyHorizStart)
            aPrevAngle = ARMDIA1.FN_CalcAngle(xyConstruct, xyCurveStart)
            nDrawnLength = (aAngle - aPrevAngle) * ARMDIA1.PI * nHeelRad / 180
            If (aAngle > aPrevAngle) Then
                aAngleInc = (aAngle - aPrevAngle) / 4
            Else
                aAngleInc = ((aAngle + 360) - aPrevAngle) / 4
            End If
            Dim ii As Double = 1
            While (ii <= 3)
                ''xyTmp = CalcXY("relpolar", xyConstruct, nHeelRad, aPrevAngle + aAngleInc * ii) ''//*
                ARMDIA1.PR_CalcPolar(xyConstruct, aPrevAngle + aAngleInc * ii, nHeelRad, xyTmp)
                If (xyTmp.x < xyAnkle.X) Then
                    ''AddVertex(xyTmp) ''//*
                    ptCurveColl.Add(New Point3d(xyTmp.X, xyTmp.y, 0))
                    xyEnd = xyTmp
                End If
                ii = ii + 1
            End While
            'AddVertex(xyHorizStart) ;
            'AddVertex(xyHorizStart.X + 0.25, xyHorizStart.y) ;
            'AddVertex(xyHorizStart.X + 0.5, xyHorizStart.y) ;
            ptCurveColl.Add(New Point3d(xyHorizStart.X, xyHorizStart.y, 0))
            ptCurveColl.Add(New Point3d(xyHorizStart.X + 0.25, xyHorizStart.y, 0))
            ptCurveColl.Add(New Point3d(xyHorizStart.X + 0.5, xyHorizStart.y, 0))

            ''// Requested zipper length
            If (nZipLength > 0) Then
                xyHorizEnd.X = xyHorizStart.X + ((nZipLength / 1.2) - nDrawnLength)
                xyHorizEnd.y = xyHorizStart.y
            Else
                nDrawnLength = nDrawnLength + ARMDIA1.FN_CalcLength(xyHorizStart, xyHorizEnd)
                nZipLength = (nDrawnLength + nElastic) * 1.2
            End If
            ''AddVertex ( xyHorizEnd )
            ptCurveColl.Add(New Point3d(xyHorizEnd.X, xyHorizEnd.y, 0))
            Dim nCurveLength = 0
            PR_DrawZipperSpline(ptCurveColl, nCurveLength)
            SetDBData("ID", sZipperID)
            SetDBData("Zipper", "1")
            SetDBData("ZipperLength", Str(nCurveLength))

            PR_DrawClosedArrow(xyCurveStart, aPrevAngle + 92)
            SetDBData("ID", sZipperID)
            SetDBData("Zipper", "1")
        Else
            If (nZipLength > 0) Then
                xyHorizEnd.X = xyHorizStart.X + ((nZipLength / 1.2))
                xyHorizEnd.y = xyHorizStart.y
            Else
                nDrawnLength = ARMDIA1.FN_CalcLength(xyHorizStart, xyHorizEnd)
                nZipLength = (nDrawnLength + nElastic) * 1.2
            End If
            PR_DrawLine(xyHorizStart, xyHorizEnd)
            SetDBData("ID", sZipperID)
            SetDBData("Zipper", "1")
            PR_DrawClosedArrow(xyHorizStart, 0)
            SetDBData("ID", sZipperID)
            SetDBData("Zipper", "1")
        End If

        ''// Add label and arrows
        If (xyHorizEnd.y > xyCO_WaistBott.y And xyCO_WaistBott.y <> 0 And xyCO_WaistBott.X <> 0 And nMedial = 0 And EOS = True) Then
            Dim xyArw As ARMDIA1.XY
            ARMDIA1.PR_MakeXY(xyArw, xyHorizEnd.X, xyCO_WaistBott.y - 0.03125)
            PR_DrawClosedArrow(xyArw, 225)
        Else
            PR_DrawClosedArrow(xyHorizEnd, 180)
        End If
        SetDBData("ID", sZipperID)
        SetDBData("Zipper", "1")
        Dim sZipperText As String = ""
        If (nMedial = 1) Then
            sZipperText = " MEDIAL ZIPPER"
        Else
            sZipperText = " LATERAL ZIPPER"
        End If
        Dim sText As String = Trim(LGLEGDIA1.fnInchesToText(nZipLength)) + Chr(34) + sZipperText
        If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
            sText = fnDisplayToCM(nZipLength) + sZipperText
        End If
        ''idLastCreated = idTextAsSymbol
        PR_DrawTextAsSymbol(xyBlock, sZipperID, sText, 0)
        'Dim idTextAsSymbol As ObjectId = idLastCreated
        SetDBData("ID", sZipperID)
        SetDBData("Zipper", "1")
        'SetDBData("Data", "INVALID ZIPPER CALCULATION")
        SetDBData("Data", sText)

        ''// Reset and exit
        ''Execute("menu", "SetLayer", Table("find", "layer", "1")) ;
        MsgBox("Zipper drawing Complete", 0, "Ankle Zipper")
    End Sub
    Public Sub PR_DrawWaistChapAnkleZipper()
        Dim xyOtemplate, xyHeel, xyAnkle, xyFold, xyEOS, xySmallestCir, xyConstruct, xyInt, xyTmp, xyEnd As ARMDIA1.XY
        Dim sUnits As String = ""
        Dim sTemplate As String = ""
        Dim sType As String = ""
        Dim sStyleID As String = ""
        Dim sLeg As String = ""
        Dim SmallHeel As Boolean = False
        Dim nX, nY As Double
        Dim EOS As Boolean
        Dim nMedial, nZipLength, nHeelRad, nHeelOff, nZipOff, nElastic, nFoldOffset, nTemplate As Double
        Dim sZipLength As String = ""
        Dim sElastic As String = ""
        Dim xyCurveStart, xyHorizStart, xyHorizEnd, xyZipperStartInChap, xyZipperEndInChap As ARMDIA1.XY
        Dim DrawCurve As Boolean = False
        Dim aAngle, aPrevAngle, aAngleInc, nDrawnLength, nZipperInChap As Double


        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim ptCurveColl As Point3dCollection = New Point3dCollection
        Dim ptEntOpts As PromptEntityOptions = New PromptEntityOptions(Chr(10) & "Select a Leg Profile")
        ptEntOpts.AllowNone = True
        ptEntOpts.SetRejectMessage(Chr(10) & "Only Curve")
        ptEntOpts.AddAllowedClass(GetType(Spline), True)
        Dim ptEntRes As PromptEntityResult = ed.GetEntity(ptEntOpts)
        If ptEntRes.Status <> PromptStatus.OK Then
            MsgBox("User cancelled", 0, "Waist Chap Ankle Zipper")
            Exit Sub
        End If

        Using tr As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim idSpline As ObjectId = ptEntRes.ObjectId
            Dim acEnt As Spline = tr.GetObject(idSpline, OpenMode.ForRead)
            Dim rb As ResultBuffer = acEnt.GetXDataForApplication("ID")
            Dim strProfile As String = ""
            If Not IsNothing(rb) Then
                ' Get the values in the xdata
                For Each typeVal As TypedValue In rb
                    If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                        Continue For
                    End If
                    strProfile = strProfile & typeVal.Value
                Next
            End If
            ''// Check if this is a valid Leg Curve exit if not
            strProfile = Trim(strProfile)
            If strProfile.Contains("LeftLegCurve") Then
                sStyleID = Mid(strProfile, 1, strProfile.Length - 12)
                sLeg = "Left"
            ElseIf strProfile.Contains("RightLegCurve") Then
                sStyleID = Mid(strProfile, 1, strProfile.Length - 13)
                sLeg = "Right"
            Else
                MsgBox("A Leg Profile was not selected" & Chr(10), 16, "Ankle Zipper")
                Exit Sub
            End If

            ''// Get Marker data
            sStyleID = sStyleID + sLeg
            Dim sHeelID As String = sStyleID + "Heel"
            Dim sOtemplateID As String = sStyleID + "Origin"
            Dim sAnkleID As String = sStyleID + "Ankle"
            Dim sFoldID As String = sStyleID + "Fold"
            Dim sEOSID As String = sStyleID + "EOS"
            Dim filterType(5) As TypedValue
            filterType.SetValue(New TypedValue(DxfCode.BlockName, "X Marker"), 0)
            filterType.SetValue(New TypedValue(DxfCode.Operator, "<or"), 1)
            filterType.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "OriginID"), 2)
            filterType.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "HeelID"), 3)
            filterType.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "ID"), 4)
            filterType.SetValue(New TypedValue(DxfCode.Operator, "or>"), 5)
            Dim selFilter As SelectionFilter = New SelectionFilter(filterType)
            Dim selResult As PromptSelectionResult = ed.SelectAll(selFilter)
            If selResult.Status <> PromptStatus.OK Then
                MsgBox("Can't find Marker Details", 16, "Ankle Zipper")
                Exit Sub
            End If
            Dim selectionSet As SelectionSet = selResult.Value
            Dim nMarkersFound As Double = 0
            Dim nMarkersRequired As Double = 5
            For Each idObject As ObjectId In selectionSet.GetObjectIds()
                Dim blkXMarker As BlockReference = tr.GetObject(idObject, OpenMode.ForRead)
                Dim position As Point3d = blkXMarker.Position
                If position.X = 0 And position.Y = 0 Then
                    Continue For
                End If
                Dim resbuf As ResultBuffer = blkXMarker.GetXDataForApplication("OriginID")
                Dim strXMarkerData As String = ""
                If Not IsNothing(resbuf) Then
                    For Each typeVal As TypedValue In resbuf
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        strXMarkerData = strXMarkerData & typeVal.Value
                    Next
                End If
                If strXMarkerData.Equals(sOtemplateID) Then
                    nMarkersFound = nMarkersFound + 1
                    ARMDIA1.PR_MakeXY(xyOtemplate, position.X, position.Y)
                    Dim resbufUnits As ResultBuffer = blkXMarker.GetXDataForApplication("units")
                    sUnits = ""
                    If Not IsNothing(resbufUnits) Then
                        For Each typeVal As TypedValue In resbufUnits
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            sUnits = sUnits & typeVal.Value
                        Next
                    End If
                    Dim resbufPressure As ResultBuffer = blkXMarker.GetXDataForApplication("Pressure")
                    sTemplate = ""
                    If Not IsNothing(resbufPressure) Then
                        For Each typeVal As TypedValue In resbufPressure
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            sTemplate = sTemplate & typeVal.Value
                        Next
                    End If
                    Dim resbufData As ResultBuffer = blkXMarker.GetXDataForApplication("Data")
                    sType = ""
                    If Not IsNothing(resbufData) Then
                        For Each typeVal As TypedValue In resbufData
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            sType = sType & typeVal.Value
                        Next
                    End If
                End If
                Dim rbHeelID As ResultBuffer = blkXMarker.GetXDataForApplication("HeelID")
                Dim strHeelID As String = ""
                If Not IsNothing(rbHeelID) Then
                    For Each typeVal As TypedValue In rbHeelID
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        strHeelID = strHeelID & typeVal.Value
                    Next
                End If
                If strHeelID.Equals(sHeelID) Then
                    nMarkersFound = nMarkersFound + 1
                    ARMDIA1.PR_MakeXY(xyHeel, position.X, position.Y)
                    Dim rbHeelData As ResultBuffer = blkXMarker.GetXDataForApplication("Data")
                    strHeelID = ""
                    If Not IsNothing(rbHeelData) Then
                        For Each typeVal As TypedValue In rbHeelData
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            strHeelID = strHeelID & typeVal.Value
                        Next
                    End If
                    Dim nSHeel As Integer = Val(strHeelID)
                    If nSHeel = 1 Then
                        SmallHeel = True
                    End If
                End If
                Dim rbAnkleID As ResultBuffer = blkXMarker.GetXDataForApplication("ID")
                Dim strAnkleID As String = ""
                If Not IsNothing(rbAnkleID) Then
                    For Each typeVal As TypedValue In rbAnkleID
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        strAnkleID = strAnkleID & typeVal.Value
                    Next
                End If
                If strAnkleID.Equals(sAnkleID) Then
                    nMarkersFound = nMarkersFound + 1
                    ARMDIA1.PR_MakeXY(xyAnkle, position.X, position.Y)
                    Dim rbAnkleData As ResultBuffer = blkXMarker.GetXDataForApplication("Data")
                    strAnkleID = ""
                    If Not IsNothing(rbAnkleData) Then
                        For Each typeVal As TypedValue In rbAnkleData
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            strAnkleID = strAnkleID & typeVal.Value
                        Next
                        If strAnkleID = "" Then
                            MsgBox("Can't get data from Ankle Marker!", 16, "Waist Chap Ankle Zipper")
                            Exit Sub
                        End If
                        ScanLine(strAnkleID, nX, nY)
                    End If
                End If
                If strAnkleID.Equals(sFoldID) Then
                    nMarkersFound = nMarkersFound + 1
                    ARMDIA1.PR_MakeXY(xyFold, position.X, position.Y)
                End If
                If strAnkleID.Equals(sEOSID) Then
                    nMarkersFound = nMarkersFound + 1
                    ARMDIA1.PR_MakeXY(xyEOS, position.X, position.Y)
                End If
            Next
            ''// Check if that the markers have been found, otherwise exit
            If sType.Contains("CHAP") = False Then
                MsgBox("Can't use this zipper for a Waist Ht style!", 16, "Waist Chap Ankle Zipper")
                Exit Sub
            End If
            If (nMarkersFound < nMarkersRequired) Then
                MsgBox("Missing markers for selected leg profile, data not found!" & Chr(10) & "Or the leg does not have a foot!", 16, "Waist Chap Ankle Zipper")
                Exit Sub
            End If
            If (nMarkersFound > nMarkersRequired) Then
                MsgBox("Two or more drawings of the same style exist!" & Chr(10) & "Delete the extra drawing/s and try again.", 16, "Ankle Zipper")
                Exit Sub
            End If
            ''// Get Smallest circumference (by requesting) a point
            ''// Do some checking that the given point is sensible
            Dim bIsLoop As Boolean = True
            While (bIsLoop)
                Dim pPtRes As PromptPointResult
                Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                pPtOpts.Message = vbLf & "Select smallest leg circumference"
                pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                If pPtRes.Status = PromptStatus.Cancel Then
                    MsgBox("User Cancelled", 16, "Ankle Zipper")
                    Exit Sub
                End If
                Dim ptStart As Point3d = pPtRes.Value
                ARMDIA1.PR_MakeXY(xySmallestCir, ptStart.X, ptStart.Y)
                If xySmallestCir.y > xyAnkle.y Or xySmallestCir.y < xyOtemplate.y Then
                    MsgBox("Selected point above Ankle or below edge of template. Try again", 16, "Waist Chap Ankle Zipper")
                Else
                    bIsLoop = False
                End If
            End While

            ''// Create Dialog
            ''// Get Zipper style
            ''//
            sUnits = "Inches"
            EOS = False
            bIsLoop = True
            Dim bIsSelect As Boolean = False
            '         sDlgElasticList = "3/8\" Elastic\n3/4\" Elastic\n1½\" Elastic\nNo Elastic";
            'sDlgElasticList =  "3/4\" Elastic\n" + sDlgElasticList;
            'sDlgLengthList = "EOS\nGive a length\nSelected Point\n";
            nMedial = 0
            While (bIsLoop)
                '               nButX = 60; nButY = 45;
                '       hDlg = Open("dialog", sLeg + "Ankle Zipper (CHAP)", "font Helv 8", 20, 20, 215, 70);

                'AddControl(hDlg, "pushbutton", nButX, nButY, 35, 14, "Cancel", "%cancel", "");
                'AddControl(hDlg, "pushbutton", nButX + 48, nButY, 35, 14, "OK", "%ok", "");

                'AddControl(hDlg, "ltext", 5, 12, 30, 14, "Ankle to", "string", "");
                'AddControl(hDlg, "combobox", 40, 10, 65, 45, sDlgLengthList, "string", "sZipLength");
                'AddControl(hDlg, "ltext", 113, 12, 33, 14, "Proximal:", "string", "");
                '	AddControl(hDlg, "dropdownlist", 145, 10, 60, 60, sDlgElasticList, "string", "sElastic");
                ' 	AddControl(hDlg, "ltext", 5, nButY - 15, 60, 15, "Template: " + sTemplate, "String", "");
                ' 	AddControl(hDlg, "checkbox", 120, nButY - 18, 60, 15, "Medial Zipper", "number", "nMedial");

                '  	Ok = Display("dialog", hDlg, "%center");
                ' 	Close("dialog", hDlg);

                'If (Ok == %cancel ) Exit (%ok, "User Cancel!") ;
                Dim oWHZipAnkFrm As New WHZIPANK_frm
                oWHZipAnkFrm.g_sCaption = sLeg + "Ankle Zipper (CHAP)"
                oWHZipAnkFrm.g_sTemplate = ""
                If sTemplate <> "" Then
                    oWHZipAnkFrm.g_sTemplate = "Template: " + sTemplate
                End If
                oWHZipAnkFrm.ShowDialog()
                If oWHZipAnkFrm.g_bIsCancel = True Then
                    MsgBox("User Cancel!", 16, "Waist Chap Ankle Zipper")
                    Exit Sub
                End If
                sZipLength = oWHZipAnkFrm.g_sZipLength
                sElastic = oWHZipAnkFrm.g_sElastic
                If oWHZipAnkFrm.g_bIsMedialZipper = True Then
                    nMedial = 1
                End If

                If (sZipLength.Equals("EOS") Or sZipLength.Equals("Selected Point")) Then
                    bIsLoop = False
                    If (sZipLength.Equals("EOS")) Then
                        EOS = True
                    Else
                        bIsSelect = True
                    End If
                Else
                    nZipLength = Val(sZipLength)
                    If (nZipLength = 0 And sZipLength.Length > 0) Then
                        bIsLoop = True
                    Else
                        bIsLoop = False
                        EOS = False
                        bIsSelect = False
                    End If
                End If
            End While

            ''// Draw Chosen Zipper style
            ''// 
            If (SmallHeel) Then
                nHeelRad = 1.1760709
                nHeelOff = 0.25
            Else
                nHeelRad = 1.5053069
                nHeelOff = 0.5
            End If
            Dim nInchToCM As Double = 1
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                nInchToCM = 2.54
            End If
            nZipOff = 1.125 * nInchToCM

            xyConstruct.x = xyAnkle.X + nX  '// nX And nY stored relative position Of the 
            xyConstruct.y = xyAnkle.y + nY  ''// heel construction circle center point 

            Dim xyCir1, xyCir2 As ARMDIA1.XY
            ARMDIA1.PR_MakeXY(xyCir1, xyHeel.X + nHeelOff, xyConstruct.y)
            ARMDIA1.PR_MakeXY(xyCir2, xyHeel.X + nHeelOff, 0)
            Dim Ok As Short = FN_CirLinInt(xyCir1, xyCir2, xyConstruct, nHeelRad + nZipOff, xyInt)
            If (Ok) Then
                ''// Draw polyline curve from intersection
                xyCurveStart = xyInt
                xyHorizStart.x = xySmallestCir.X
                xyHorizStart.y = xySmallestCir.y - nZipOff
                DrawCurve = True
            Else
                ''// Otherwise degenerate to a straight line
                xyHorizStart.x = xyHeel.X + nHeelOff
                xyHorizStart.y = xyHeel.y - nZipOff
                DrawCurve = False
            End If
            nZipperInChap = 0
            ''//if  EOS given
            If (EOS) Then
                If (nMedial = 1) Then
                    xyHorizEnd.x = xyFold.X
                    nElastic = 0
                    nZipperInChap = 0
                Else
                    xyZipperStartInChap.x = xyFold.X
                    xyZipperStartInChap.y = xyFold.y - 1.125
                    xyZipperEndInChap = xyEOS
                    xyZipperEndInChap.y = xyFold.y - 1.125
                    nZipperInChap = ARMDIA1.FN_CalcLength(xyZipperStartInChap, xyZipperEndInChap)
                    xyHorizEnd.x = xyFold.X
                    ''// Elastic allowance at EOS
                    nElastic = 0.75
                    If (String.Compare(Mid(sElastic, 1, 1), "N") = 0) Then
                        nElastic = 0
                    End If
                    If (String.Compare(Mid(sElastic, 1, 3), "3/8") = 0) Then
                        nElastic = 0.375
                    End If
                    If (String.Compare(Mid(sElastic, 1, 1), "1") = 0) Then
                        nElastic = 1.5
                    End If
                End If
                xyHorizEnd.y = xyHorizStart.y
            Else
                ''// No elastic as this implies a closed zipper
                nElastic = 0
                ''//create a dummy xyHorizEnd for use on insertion
                If (xyHorizStart.X + (nZipLength / 1.2) > xyFold.X) Then
                    xyHorizEnd.X = xyFold.X
                Else
                    xyHorizEnd.X = xyHorizStart.X + (nZipLength / 1.2)
                    xyHorizEnd.y = xyHorizStart.y
                End If
            End If

            ''// For Medial Zipper step back 3 tape
            If (nMedial = 1) Then
                nFoldOffset = 0
                nTemplate = Val(Mid(sTemplate, 1, 2))
                If (nTemplate >= 30) Then nFoldOffset = 1.25 * 3
                If (nTemplate = 13) Then
                    nFoldOffset = 1.31 * 3
                    MsgBox("WARNING!" & Chr(10) & "Template drawn using 13 Down Stretch!", 16, "Waist Chap Ankle Zipper")
                End If
                If (nTemplate = 9) Then nFoldOffset = 1.37 * 3
                xyHorizEnd.X = xyHorizEnd.X - nFoldOffset
            End If

            ''// If select given then prompt user for end of zip
            ''//
            If (bIsSelect) Then
                Dim pEndPtRes As PromptPointResult
                Dim pEndPtOpts As PromptPointOptions = New PromptPointOptions("")
                pEndPtOpts.Message = vbLf & "Select End of Zipper"
                pEndPtRes = acDoc.Editor.GetPoint(pEndPtOpts)
                If pEndPtRes.Status = PromptStatus.Cancel Then
                    MsgBox("End not selected", 16, "Waist Chap Ankle Zipper")
                    Exit Sub
                End If
                Dim ptEndZip As Point3d = pEndPtRes.Value
                ARMDIA1.PR_MakeXY(xyHorizEnd, ptEndZip.X, ptEndZip.Y)
                xyHorizEnd.y = xyHorizStart.y
                If (xyHorizEnd.X > xyFold.X) Then
                    xyZipperStartInChap.X = xyFold.X
                    xyZipperStartInChap.y = xyFold.y - 1.125
                    xyZipperEndInChap.X = xyHorizEnd.X
                    xyZipperEndInChap.y = xyFold.y - 1.125
                    If (xyZipperEndInChap.X < xyFold.X) Then
                        nZipperInChap = 0
                    Else
                        nZipperInChap = ARMDIA1.FN_CalcLength(xyZipperStartInChap, xyZipperEndInChap)
                    End If
                    xyHorizEnd.X = xyFold.X
                    nElastic = 0
                    EOS = True
                End If
            End If
        End Using
        ''// Revise radius to include Zip offset
        ''//
        nHeelRad = nHeelRad + nZipOff
        ARMDIA1.PR_SetLayer("Notes")
        Dim xyBlock As ARMDIA1.XY
        ARMDIA1.PR_MakeXY(xyBlock, xyHorizStart.X + (xyHorizEnd.X - xyHorizStart.X) / 2, xyHorizStart.y + 0.125)
        ''PR_DrawTextAsSymbol(xyBlock, "", "", 0)
        ''// Create ID string
        Dim sZipperID As String = sStyleID ''+ MakeString("scalar", UID("get", hTextEnt))
        ''SetDBData("ID", sZipperID)
        ''SetDBData("Zipper", "1")
        ''SetDBData("Data", "INVALID ZIPPER CALCULATION") ''//Do this just incase it fails
        If (DrawCurve) Then
            ''StartPoly("fitted");
            ''AddVertex(xyCurveStart);
            ptCurveColl.Add(New Point3d(xyCurveStart.X, xyCurveStart.y, 0))
            aAngle = ARMDIA1.FN_CalcAngle(xyConstruct, xyHorizStart)
            aPrevAngle = ARMDIA1.FN_CalcAngle(xyConstruct, xyCurveStart)
            If (aAngle > aPrevAngle) Then
                aAngleInc = (aAngle - aPrevAngle) / 4
            Else
                aAngleInc = ((aAngle + 360) - aPrevAngle) / 4
            End If
            Dim ii As Double = 1
            While (ii <= 3)     ''//*
                ''xyTmp = CalcXY("relpolar", xyConstruct, nHeelRad, aPrevAngle + aAngleInc * ii) ''//*
                ARMDIA1.PR_CalcPolar(xyConstruct, aPrevAngle + aAngleInc * ii, nHeelRad, xyTmp)
                If (xyTmp.x < xyAnkle.X) Then
                    ''AddVertex(xyTmp) ''; //*
                    ptCurveColl.Add(New Point3d(xyTmp.X, xyTmp.y, 0))
                    xyEnd = xyTmp
                End If
                ii = ii + 1
            End While
            ''AddVertex(xyHorizStart) ;
            ''AddVertex(xyHorizStart.X + 0.25, xyHorizStart.y) ;
            ''AddVertex(xyHorizStart.X + 0.5, xyHorizStart.y) ;		// Note this 0.5 Is added below For nZipLength > 0
            ''EndPoly();	
            ptCurveColl.Add(New Point3d(xyHorizStart.X, xyHorizStart.y, 0))
            ptCurveColl.Add(New Point3d(xyHorizStart.X + 0.25, xyHorizStart.y, 0))
            ptCurveColl.Add(New Point3d(xyHorizStart.X + 0.5, xyHorizStart.y, 0))
            Dim nCurveLength As Double = 0
            PR_DrawZipperSpline(ptCurveColl, nCurveLength)
            ''hEnt = UID("find", UID("getmax")) ;	
            SetDBData("ID", sZipperID)
            SetDBData("Zipper", "1")
            SetDBData("ZipperLength", Str(nCurveLength))
            ''------GetDBValue(hEnt, "ZipperLength", & sTmp, & nDrawnLength);
            nDrawnLength = nCurveLength
            nZipperInChap = 0
            ''// Requested zipper length
            If (nZipLength > 0) Then
                xyHorizEnd.X = xyHorizStart.X + 0.5 + ((nZipLength / 1.2) - nDrawnLength)
                xyHorizEnd.y = xyHorizStart.y
                If (xyHorizEnd.X > xyFold.X) Then
                    ''//Draw the bit in the chap
                    nZipperInChap = xyHorizEnd.X - xyFold.X
                    xyZipperStartInChap.X = xyFold.X
                    xyZipperStartInChap.y = xyFold.y - 1.125
                    xyZipperEndInChap.X = xyHorizEnd.X
                    xyZipperEndInChap.y = xyFold.y - 1.125
                    xyHorizEnd.X = xyFold.X
                    nElastic = 0
                End If
                ''SetVertex(hEnt, nVertex, xyHorizEnd)
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    Dim acSpline As Spline = acTrans.GetObject(idLastCreated, OpenMode.ForWrite)
                    If acSpline.IsNull <> True Then
                        ''Dim nLen As Double = System.Math.Abs(xyFirstProfile.y - xyOtemplate.y)
                        ''xyFirstProfile.y = xyOtemplate.y + (nLen * nInchToCM)
                        ''acSpline.SetFitPointAt(ptCurveColl.Count + 1, New Point3d(xyHorizEnd.X, xyHorizEnd.y, 0))
                        acSpline.Erase()
                    End If
                    acTrans.Commit()
                End Using
                ptCurveColl.Add(New Point3d(xyHorizEnd.X, xyHorizEnd.y, 0))
                PR_DrawZipperSpline(ptCurveColl, nCurveLength)
            Else
                ''SetVertex(hEnt, nVertex, xyHorizEnd)
                Using acTransaction As Transaction = acCurDb.TransactionManager.StartTransaction()
                    Dim acSpline As Spline = acTransaction.GetObject(idLastCreated, OpenMode.ForWrite)
                    If acSpline.IsNull <> True Then
                        ''Dim nLen As Double = System.Math.Abs(xyFirstProfile.y - xyOtemplate.y)
                        ''xyFirstProfile.y = xyOtemplate.y + (nLen * nInchToCM)
                        ''acSpline.SetFitPointAt(ptCurveColl.Count + 1, New Point3d(xyHorizEnd.X, xyHorizEnd.y, 0))
                        acSpline.Erase()
                    End If
                    acTransaction.Commit()
                End Using
                ptCurveColl.Add(New Point3d(xyHorizEnd.X, xyHorizEnd.y, 0))
                PR_DrawZipperSpline(ptCurveColl, nCurveLength)
                ''--------GetDBValue(hEnt, "ZipperLength", & sTmp, & nDrawnLength)
                nDrawnLength = nCurveLength
                nZipLength = (nZipperInChap + nDrawnLength + nElastic) * 1.2
            End If
            PR_DrawClosedArrow(xyCurveStart, aPrevAngle + 92)
            SetDBData("ID", sZipperID)
            SetDBData("Zipper", "1")
        Else
            If (nZipLength > 0) Then
                xyHorizEnd.X = xyHorizStart.X + ((nZipLength / 1.2))
                xyHorizEnd.y = xyHorizStart.y
            Else
                nDrawnLength = ARMDIA1.FN_CalcLength(xyHorizStart, xyHorizEnd)
                nZipLength = (nDrawnLength + nElastic) * 1.2
            End If
            PR_DrawLine(xyHorizStart, xyHorizEnd)
            SetDBData("ID", sZipperID)
            SetDBData("Zipper", "1")
            PR_DrawClosedArrow(xyHorizStart, 0)
            SetDBData("ID", sZipperID)
            SetDBData("Zipper", "1")
        End If

        ''// Add label and arrows

        ''// If a zipper extends to the EOS then we draw this bit
        If (nZipperInChap <> 0) Then
            PR_DrawLine(xyZipperStartInChap, xyZipperEndInChap)
            SetDBData("ID", sZipperID)
            SetDBData("Zipper", "1")
            PR_DrawClosedArrow(xyZipperStartInChap, 0)
            SetDBData("ID", sZipperID)
            SetDBData("Zipper", "1")
            PR_DrawClosedArrow(xyZipperEndInChap, 180)
            SetDBData("ID", sZipperID)
            SetDBData("Zipper", "1")
            ''// Get Velcro overlap (If it exists!) And delete it.
            ''-------sTmp = "DB ID ='" + sStyleID + "' AND DB Zipper ='VelcroOverlap'";
            ''-----------hChan = Open("selection", sTmp);
            Dim filterVelcro(1) As TypedValue
            filterVelcro.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "ID"), 0)
            filterVelcro.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "Zipper"), 1)
            Dim selVelcroFilter As SelectionFilter = New SelectionFilter(filterVelcro)
            Dim selVelcroResult As PromptSelectionResult = ed.SelectAll(selVelcroFilter)
            If selVelcroResult.Status <> PromptStatus.OK Then
                MsgBox("Can't find Marker Details", 16, "Ankle Zipper")
                Exit Sub
            End If
            Dim selectionSetVelcro As SelectionSet = selVelcroResult.Value
            Using acTr As Transaction = acCurDb.TransactionManager.StartTransaction()
                For Each idObject As ObjectId In selectionSetVelcro.GetObjectIds()
                    Dim acVelcroEnt As Entity = acTr.GetObject(idObject, OpenMode.ForWrite)
                    Dim rbID As ResultBuffer = acVelcroEnt.GetXDataForApplication("ID")
                    Dim strID As String = ""
                    If Not IsNothing(rbID) Then
                        ' Get the values in the xdata
                        For Each typeVal As TypedValue In rbID
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            strID = strID & typeVal.Value
                        Next
                    End If
                    If strID.Contains(sStyleID) = False Then
                        Continue For
                    End If

                    Dim rbZipper As ResultBuffer = acVelcroEnt.GetXDataForApplication("Zipper")
                    Dim strZipper As String = ""
                    If Not IsNothing(rbZipper) Then
                        ' Get the values in the xdata
                        For Each typeVal As TypedValue In rbZipper
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            strZipper = strZipper & typeVal.Value
                        Next
                    End If
                    If strZipper.Contains("VelcroOverlap") = False Then
                        Continue For
                    End If
                    Dim sLayer As String = acVelcroEnt.Layer
                    If (sLayer.ToUpper().Equals("NOTES") And acVelcroEnt.GetType().Name.ToUpper.Equals("LINE")) Then
                        Dim acLyrTbl As LayerTable
                        acLyrTbl = acTr.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                        If acLyrTbl.Has("Template" + sLeg) = False Then
                            Continue For
                        End If
                        acVelcroEnt.SetLayerId(acLyrTbl("Template" + sLeg), True)
                        acVelcroEnt.XData = New ResultBuffer(New TypedValue(DxfCode.ExtendedDataRegAppName, "Zipper"))
                    Else
                        acVelcroEnt.Erase()
                    End If
                Next
                acTr.Commit()
            End Using
        Else
            PR_DrawClosedArrow(xyHorizEnd, 180)
            SetDBData("ID", sZipperID)
            SetDBData("Zipper", "1")
        End If

        ''// Update text symbol with Zipper Length
        Dim sZipperText As String = ""
        If (nMedial = 1) Then
            sZipperText = " MEDIAL ZIPPER"
        Else
            sZipperText = " LATERAL ZIPPER"
        End If
        Dim sText As String = Trim(LGLEGDIA1.fnInchesToText(nZipLength)) + Chr(34) + sZipperText
        If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
            sText = fnDisplayToCM(nZipLength) + sZipperText
        End If
        PR_DrawTextAsSymbol(xyBlock, sZipperID, sText, 0)
        SetDBData("ID", sZipperID)
        SetDBData("Zipper", "1")
        SetDBData("Data", sText)

        ''// Reset And exit
        ''Execute("menu", "SetLayer", Table("find", "layer", "1")) ;
        MsgBox("Zipper drawing Complete", 0, "Waist Chap Ankle Zipper")
    End Sub
    Public Sub PR_DrawGlovePalmZipper()
        Dim xyPALMER, xyPALMERWEB, xyStartEOS, xyEndEOS, xyTmp, xyZipStart, xyZipEnd As ARMDIA1.XY
        Dim strProfile As String = ""
        Dim sSide As String = ""
        Dim nDirection As Integer
        Dim sAge As String = ""
        Dim sData As String = ""
        Dim nType, nAge, nZipLength, nElasticProximal, nElastic, nElasticFactor, nWebOffSet As Double
        Dim nInsertSize As Integer
        Dim sZipLength As String = ""
        Dim sElasticProximal As String = ""
        Dim sWebOffSet As String = ""
        Dim EOStoSelectedPoint, EOStoCalculatedPoint, bIsLoop As Boolean
        Dim xyPt1, xyPt2, xyGivenPoint As ARMDIA1.XY
        Dim nDrawnLength As Double

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim ptEntOpts As PromptEntityOptions = New PromptEntityOptions("Select Glove Profile")
        ptEntOpts.AllowNone = True
        ''ptEntOpts.Keywords.Add("Spline")
        Dim ptEntRes As PromptEntityResult = ed.GetEntity(ptEntOpts)
        If ptEntRes.Status <> PromptStatus.OK Then
            MsgBox("User Cancelled", 16, "Palm Zipper")
            Exit Sub
        End If

        Dim idEnt As ObjectId = ptEntRes.ObjectId
        Using tr As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acEnt As Entity = tr.GetObject(idEnt, OpenMode.ForRead)
            Dim rb As ResultBuffer = acEnt.GetXDataForApplication("ID")
            If Not IsNothing(rb) Then
                ' Get the values in the xdata
                For Each typeVal As TypedValue In rb
                    If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                        Continue For
                    End If
                    strProfile = strProfile & typeVal.Value
                Next
            End If
            If strProfile.Contains("Left") Then
                sSide = "Left"
                nDirection = 1
            ElseIf strProfile.Contains("Right") Then
                sSide = "Right"
                nDirection = -1
            Else
                MsgBox("A Glove Profile was not selected", 16, "Palm Zipper")
                Exit Sub
            End If

            ''// Get data for a PALMER Zipper
            Dim nFound As Integer = 0
            Dim nReqired As Integer = 3
            Dim filterType(5) As TypedValue
            filterType.SetValue(New TypedValue(DxfCode.Operator, "<or"), 0)
            filterType.SetValue(New TypedValue(DxfCode.BlockName, "X Marker"), 1)
            filterType.SetValue(New TypedValue(DxfCode.Start, "Line"), 2)
            filterType.SetValue(New TypedValue(DxfCode.Operator, "or>"), 3)
            filterType.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "ID"), 4)
            filterType.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "Zipper"), 5)
            Dim selFilter As SelectionFilter = New SelectionFilter(filterType)
            Dim selResult As PromptSelectionResult = ed.SelectAll(selFilter)
            If selResult.Status <> PromptStatus.OK Then
                MsgBox("Can't get PALMER Zipper Details", 16, "Palm Zipper")
                Exit Sub
            End If
            Dim selectionSet As SelectionSet = selResult.Value
            For Each idObject As ObjectId In selectionSet.GetObjectIds()
                Dim acEntity As Entity = tr.GetObject(idObject, OpenMode.ForRead)
                Dim rbID As ResultBuffer = acEntity.GetXDataForApplication("ID")
                Dim strID As String = ""
                If Not IsNothing(rbID) Then
                    ' Get the values in the xdata
                    For Each typeVal As TypedValue In rbID
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        strID = strID & typeVal.Value
                    Next
                End If
                If strID.Equals(strProfile) = False Then
                    Continue For
                End If

                ''Dim ptPosition As Point3d = blkXMarker.Position
                Dim ptPosition As New Point3d
                If (acEntity.GetType().Name.ToUpper.Equals("BLOCKREFERENCE")) Then
                    Dim blkXMarker As BlockReference = TryCast(acEntity, BlockReference)
                    ptPosition = blkXMarker.Position
                End If
                Dim rbZipper As ResultBuffer = acEntity.GetXDataForApplication("Zipper")
                Dim strZipper As String = ""
                If Not IsNothing(rbZipper) Then
                    ' Get the values in the xdata
                    For Each typeVal As TypedValue In rbZipper
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        strZipper = strZipper & typeVal.Value
                    Next
                End If
                If strZipper.Equals("PALMER") Then
                    nFound = nFound + 1
                    ARMDIA1.PR_MakeXY(xyPALMER, ptPosition.X, ptPosition.Y)
                    Dim rbAge As ResultBuffer = acEntity.GetXDataForApplication("Age")
                    sAge = ""
                    If Not IsNothing(rbAge) Then
                        ' Get the values in the xdata
                        For Each typeVal As TypedValue In rbAge
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            sAge = sAge & typeVal.Value
                        Next
                    End If
                    Dim rbData As ResultBuffer = acEntity.GetXDataForApplication("Data")
                    sData = ""
                    If Not IsNothing(rbData) Then
                        ' Get the values in the xdata
                        For Each typeVal As TypedValue In rbData
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            sData = sData & typeVal.Value
                        Next
                    End If
                    nType = Val(Mid(sData, 10, 2))
                    nAge = Val(sAge)
                End If
                If strZipper.Equals("PALMER-WEB") Then
                    nFound = nFound + 1
                    ARMDIA1.PR_MakeXY(xyPALMERWEB, ptPosition.X, ptPosition.Y)
                    ''// Ignore insert size for Palmer Zippers
                    nInsertSize = 0
                End If
                If strZipper.Equals("EOS") Then
                    nFound = nFound + 1
                    ''--------hEOS = hEnt 
                    ''-------GetGeometry(hEnt, & xyStartEOS, & xyEndEOS)
                    If (acEntity.GetType().Name.ToUpper.Equals("LINE")) Then
                        Dim acLine As Line = TryCast(acEntity, Line)
                        Dim ptStart As Point3d = acLine.StartPoint
                        Dim ptEnd As Point3d = acLine.EndPoint
                        ARMDIA1.PR_MakeXY(xyStartEOS, ptStart.X, ptStart.Y)
                        ARMDIA1.PR_MakeXY(xyEndEOS, ptEnd.X, ptEnd.Y)
                    End If
                End If
            Next

            ''// Check that sufficent data have been found, otherwise exit
            ''//
            If (nFound < nReqired) Then
                MsgBox("Missing data for selected Glove!", 16, "Palm Zipper")
                Exit Sub
            End If
            If (nFound > nReqired) Then
                MsgBox("Two or more drawings of the same style exist!" & Chr(10) & "Delete the extra drawing/s and try again.", 16, "Palm Zipper")
                Exit Sub
            End If

            ''// Create Dialog
            ''// Get Zipper style
            ''//
            Dim sUnits As String = "Inches"
            EOStoSelectedPoint = False
            EOStoCalculatedPoint = False

            ''// Proximal elastic
            '   sDlgElasticList =  "3/8\" Elastic\n3/4\" Elastic\n1½\" Elastic\nNo Elastic";
            '   if (nAge < 10)
            '   	sDlgElasticList =  "3/8\" Elastic\n" + sDlgElasticList;  // 1/2" for children under 10
            '   else
            '      	sDlgElasticList =  "3/4\" Elastic\n" + sDlgElasticList;  // Inch for adults 

            '// Length specification      	
            '   sDlgLengthList =  "Standard\nGive a length\nSelected Point";

            '// Offset from web
            '   sDlgWebOffSetList =  "1-1/8\"\n3/4\"";
            '   sDlgWebOffSetList =  "1-1/8\"\n"  + sDlgWebOffSetList ;	// Set Default
            bIsLoop = True
            While (bIsLoop)
                'nButX = 65; nButY = 55;
                'hDlg = Open("dialog", sSide + " Glove Zipper (PALMER)", "font Helv 8", 20, 20, 210, 75);

                'AddControl(hDlg, "pushbutton", nButX, nButY, 35, 14, "Cancel", "%cancel", "");
                'AddControl(hDlg, "pushbutton", nButX + 48, nButY, 35, 14, "OK", "%ok", "");

                'AddControl(hDlg, "ltext", 5, 12, 28, 14, "EOS to", "string", "");
                'AddControl(hDlg, "combobox", 35, 10, 70, 40, sDlgLengthList, "string", "sZipLength");

                'AddControl(hDlg, "ltext", 110, 12, 30, 14, "Proximal:", "string", "");
                '	AddControl(hDlg, "dropdownlist", 145, 10, 60, 70, sDlgElasticList, "string", "sElasticProximal");

                'AddControl(hDlg, "ltext", 100, 32, 50, 14, "Below Web:", "string", "");
                '	AddControl(hDlg, "dropdownlist", 145, 30, 60, 40, sDlgWebOffSetList, "string", "sWebOffSet");

                '     	Ok = Display("dialog", hDlg, "%center");
                ' 	Close("dialog", hDlg);

                'If (Ok == %cancel ) Exit (%ok, "User Cancel!") ;
                Dim oZipperPalmFrm As New ZIPPALM_frm
                oZipperPalmFrm.g_sCaption = sSide + " Glove Zipper (PALMER)"
                oZipperPalmFrm.g_nAge = nAge
                oZipperPalmFrm.ShowDialog()
                If oZipperPalmFrm.g_bIsCancel = True Then
                    MsgBox("User Cancel!", 16, "Palm Zipper")
                    Exit Sub
                End If
                sZipLength = oZipperPalmFrm.g_sZipLength
                sElasticProximal = oZipperPalmFrm.g_sElasticProximal
                sWebOffSet = oZipperPalmFrm.g_sWebOffset

                If sZipLength.Equals("Selected Point") Then
                    EOStoSelectedPoint = True
                End If
                If sZipLength.Equals("Standard") Then
                    EOStoCalculatedPoint = True
                End If
                If (EOStoSelectedPoint = True Or EOStoCalculatedPoint = True) Then
                    bIsLoop = False
                Else
                    nZipLength = Val(sZipLength)
                    If (nZipLength = 0 And sZipLength.Length > 0) Then
                        MsgBox("Invalid given length!" & Chr(10) & Chr(10) & "To use this option, type over the text in the " &
                               Chr(34) & "EOS to: " & Chr(34) & "box with the required length in Inches (Decimal Inches).", 16, "Palm Zipper")
                        bIsLoop = True
                    Else
                        bIsLoop = False
                    End If
                End If
            End While
            ''// Establish allowance for zippers
            ''//
            nElasticProximal = 0.75
            If (String.Compare(Mid(sElasticProximal, 1, 1), "N") = 0) Then
                nElasticProximal = 0
            End If
            If (String.Compare(Mid(sElasticProximal, 1, 3), "3/8") = 0) Then
                nElasticProximal = 0.375
            End If
            If (String.Compare(Mid(sElasticProximal, 1, 1), "1") = 0) Then
                nElasticProximal = 1.5
            End If
            nElastic = nElasticProximal
            ''// Elastic factor
            ''// Set by age
            If (nAge <= 6) Then
                nElasticFactor = 1
            Else
                nElasticFactor = 0.95     ''// Glove to elbow & Normal Glove
                If (nType = 2) Then
                    nElasticFactor = 0.92   ''// Glove to axilla
                End If
            End If

            ''// Establish Minimum offset from web    
            nWebOffSet = 1.125
            If (String.Compare(Mid(sWebOffSet, 1, 1), "3") = 0) Then
                nWebOffSet = 0.75
            End If
        End Using
        ''// Draw on layer Note
        ARMDIA1.PR_SetLayer("Notes")
        ''// Draw Zipper
        ''// 
        ''// Establish EOS start And end
        ''//
        Dim nLengthToStart As Double = ARMDIA1.FN_CalcLength(xyStartEOS, xyPALMER)
        Dim nLengthToEnd As Double = ARMDIA1.FN_CalcLength(xyEndEOS, xyPALMER)
        If (nLengthToStart > nLengthToEnd) Then
            ''// Swap start and end
            xyTmp = xyStartEOS
            xyStartEOS = xyEndEOS
            xyEndEOS = xyTmp
        End If

        ''// Angles
        ''//
        Dim aEOS As Double = ARMDIA1.FN_CalcAngle(xyStartEOS, xyEndEOS)
        Dim aPALMER As Double = ARMDIA1.FN_CalcAngle(xyStartEOS, xyPALMER)

        ''// Zipper start point on EOS
        ''//
        ''Dim aAngle As Double = System.Math.Abs(aEOS - aPALMER)
        Dim aAngle As Double = aEOS - aPALMER
        Dim nLength As Double = ARMDIA1.FN_CalcLength(xyStartEOS, xyPALMER)
        Dim nB As Double = System.Math.Cos(aAngle) * nLength
        ''xyZipStart = CalcXY("relpolar", xyStartEOS, (nB + 0.625) - nInsertSize, aEOS)
        ARMDIA1.PR_CalcPolar(xyStartEOS, aEOS, ((nB + 0.625) - nInsertSize) / 4, xyZipStart)
        ''ARMDIA1.PR_CalcPolar(xyStartEOS, (nB + 0.625) - nInsertSize, aEOS, xyZipStart)

        Dim aPALMERWEB As Double = ARMDIA1.FN_CalcAngle(xyStartEOS, xyPALMERWEB)
        aAngle = System.Math.Abs(aEOS - aPALMERWEB)
        nLength = ARMDIA1.FN_CalcLength(xyStartEOS, xyPALMERWEB)
        Dim nZipConstructLen As Double = System.Math.Abs(System.Math.Sin(aAngle) * nLength)
        Dim aZipper As Double = aEOS + (90 * nDirection)

        If (EOStoCalculatedPoint = True) Then
            ''xyZipEnd = CalcXY ("relpolar", xyZipStart, nZipConstructLen  - nWebOffSet, aZipper) 
            ARMDIA1.PR_CalcPolar(xyZipStart, aZipper, nZipConstructLen - nWebOffSet, xyZipEnd)
        End If
        If (EOStoSelectedPoint = True) Then
            ''// Only allow the point if it lies within the rectangle formed by the
            ''// EOS and the xyPALMERWEB plus a .0625" tolerance
            ''// EOS = length and angle, xyPALMERWEB = width
            ''xyPt1 = CalcXY("relpolar", xyStartEOS, nZipConstructLen + 0.0625, aZipper)
            ARMDIA1.PR_CalcPolar(xyStartEOS, aZipper, nZipConstructLen + 0.0625, xyPt1)
            ''xyPt1 = CalcXY("relpolar", xyPt1, 2, (aEOS + (nDirection * 180)))
            ARMDIA1.PR_CalcPolar(xyPt1, (aEOS + (nDirection * 180)), 2, xyPt1)
            ''xyPt2 = CalcXY("relpolar", xyEndEOS, nZipConstructLen + 0.0625, aZipper)
            ARMDIA1.PR_CalcPolar(xyEndEOS, aZipper, nZipConstructLen + 0.0625, xyPt2)
            ''xyPt2 = CalcXY ("relpolar", xyPt2,  2, aEOS )
            ARMDIA1.PR_CalcPolar(xyPt2, aEOS, 2, xyPt2)

            bIsLoop = True
            While (bIsLoop)
                Dim pPtRes As PromptPointResult
                Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                pPtOpts.Message = vbLf & "Select End of Zipper "
                pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                If pPtRes.Status = PromptStatus.Cancel Then
                    MsgBox("End Not Selected", 16, "Palm Zipper")
                    Exit Sub
                End If
                Dim ptEnd As Point3d = pPtRes.Value
                ARMDIA1.PR_MakeXY(xyGivenPoint, ptEnd.X, ptEnd.Y)
                ''If (!Calc("inpoly", xyGivenPoint, xyStartEOS, xyPt1, xyPt2, xyEndEOS)) Then
                If (PR_CheckPointExistInside(xyGivenPoint, xyStartEOS, xyPt1, xyPt2, xyEndEOS) = False) Then
                    If MsgBox("Given point can't be used! Try again." & Chr(10) & "Or use cancel to Exit", 1, "Palm Zipper") = MsgBoxResult.Cancel Then
                        MsgBox("User Cancelled", 16, "Palm Zipper")
                        Exit Sub
                    End If
                Else
                    bIsLoop = False
                End If
            End While

            aPALMERWEB = ARMDIA1.FN_CalcAngle(xyStartEOS, xyGivenPoint)
            aAngle = System.Math.Abs(aEOS - aPALMERWEB)
            nLength = ARMDIA1.FN_CalcLength(xyStartEOS, xyGivenPoint)
            Dim nA As Double = System.Math.Abs(System.Math.Sin(aAngle) * nLength)
            ''xyZipEnd = CalcXY("relpolar", xyZipStart, nA, aZipper)
            ARMDIA1.PR_CalcPolar(xyZipStart, aZipper, nA, xyZipEnd)

            ''// Ensure that selected point is no closer than nWebOffSet to the web
            ''// within 1/8th of an inch
            nLength = ARMDIA1.FN_CalcLength(xyZipStart, xyZipEnd) - nZipConstructLen
            If ((System.Math.Abs(nLength) < (nWebOffSet - 0.125)) Or (nLength > 0)) Then
                Dim sText As String = ""
                If (nLength > 0) Then
                    sText = "The end of the zipper will above the nearest web"
                Else
                    sText = "The end of the zipper will be closer to the nearest web than " +
                        Format("length", nWebOffSet) + ".  Actual distance is " + Format("length", System.Math.Abs(nLength))
                End If
                Dim ok As MsgBoxResult = MsgBox(sText & Chr(10) & "Use YES to use this point or" & Chr(10) & "Use NO to default to Standard.", 3, "Palm Zipper")
                If ok = MsgBoxResult.Cancel Then
                    MsgBox("User Cancelled", 16, "Palm Zipper")
                End If
                If ok = MsgBoxResult.No Then
                    ''xyZipEnd = CalcXY ("relpolar", xyZipStart, nZipConstructLen  - nWebOffSet, aZipper)
                    ARMDIA1.PR_CalcPolar(xyZipStart, aZipper, nZipConstructLen - nWebOffSet, xyZipEnd)
                End If
            End If
        End If

        ''// Specified ZIP Length
        ''//	
        If (nZipLength > 0) Then
            ''xyZipEnd = CalcXY ("relpolar", xyZipStart, (nZipLength * nElasticFactor) - nElastic, aZipper)
            ARMDIA1.PR_CalcPolar(xyZipStart, aZipper, (nZipLength * nElasticFactor) - nElastic, xyZipEnd)
            nDrawnLength = ARMDIA1.FN_CalcLength(xyZipStart, xyZipEnd)
            nLength = nDrawnLength - nZipConstructLen
            If ((System.Math.Abs(nLength) < (nWebOffSet - 0.125)) Or (nLength > 0)) Then
                Dim status As MsgBoxResult = MsgBox("The end of the zipper will be above or closer to the nearest web than" +
                                                    Format("length", nWebOffSet) + " with the given length." +
                                                    Chr(10) + "Use OK to default to Standard or" + Chr(10) + "Use CANCEL to exit.", 1, "Palm Zipper")
                If status = MsgBoxResult.Cancel Then
                    MsgBox("User Cancelled", 16, "Palm Zipper")
                End If
                ''xyZipEnd = CalcXY ("relpolar", xyZipStart, nZipConstructLen  - nWebOffSet, aZipper) 
                ARMDIA1.PR_CalcPolar(xyZipStart, aZipper, nZipConstructLen - nWebOffSet, xyZipEnd)
                nDrawnLength = ARMDIA1.FN_CalcLength(xyZipStart, xyZipEnd)
                nZipLength = (nDrawnLength + nElastic) / nElasticFactor
            End If
        Else
            nDrawnLength = ARMDIA1.FN_CalcLength(xyZipStart, xyZipEnd)
            nZipLength = (nDrawnLength + nElastic) / nElasticFactor
        End If

        sZipLength = Format("length", nZipLength)
        ''// Draw markers And Text label
        ''// Each entity Is given an ID Data Base value that links them together
        ''// This Is based on the UID of the symbol "ZipperText" And the Glove Profile ID.
        ''// A Symbol Is used as DB values can't be attached to text entities 

        ''// Add text symbol
        ''xyTmp = CalcXY("relpolar", xyZipStart, nDrawnLength / 2, aZipper)
        ARMDIA1.PR_CalcPolar(xyZipStart, aZipper, nDrawnLength / 2, xyTmp)
        ''// Create ID string
        Dim sZipperID As String = strProfile ''+ MakeString("scalar", UID("get", hEnt))
        PR_DrawTextAsSymbol(xyTmp, sZipperID, sZipLength + " Palmer", aZipper - 180)
        ''// Label entity with ID string and make Zipper %true
        SetDBData("ID", sZipperID)
        SetDBData("Zipper", "1")
        SetDBData("Data", sZipLength + " Palmer")

        ''// Add label and arrows
        PR_DrawClosedArrow(xyZipStart, aEOS + (90 * nDirection))
        SetDBData("ID", sZipperID)
        SetDBData("Zipper", "1")

        PR_DrawClosedArrow(xyZipEnd, aEOS - (90 * nDirection))
        SetDBData("ID", sZipperID)
        SetDBData("Zipper", "1")

        PR_DrawLine(xyZipStart, xyZipEnd)
        SetDBData("ID", sZipperID)
        SetDBData("Zipper", "1")

        ''// Reset And exit
        ''Execute("menu", "SetLayer", Table("find", "layer", "1")) 
        MsgBox("Zipper drawing Complete", 0, "Palm Zipper")
    End Sub
    Private Function PR_CheckPointExistInside(ByRef xyGivenPoint As ARMDIA1.XY, ByRef xyPt1 As ARMDIA1.XY, ByRef xyPt2 As ARMDIA1.XY,
                                              ByRef xyPt3 As ARMDIA1.XY, ByRef xyPt4 As ARMDIA1.XY) As Boolean

        Dim acPoly As Polyline = New Polyline
        acPoly.AddVertexAt(0, New Point2d(xyPt1.X, xyPt1.y), 0, 0, 0)
        acPoly.AddVertexAt(1, New Point2d(xyPt2.X, xyPt2.y), 0, 0, 0)
        acPoly.AddVertexAt(2, New Point2d(xyPt3.X, xyPt3.y), 0, 0, 0)
        acPoly.AddVertexAt(3, New Point2d(xyPt4.X, xyPt4.y), 0, 0, 0)

        Dim extBoundary As Extents3d
        'extBoundary.AddPoint(New Point3d(xyPt1.X, xyPt1.y, 0))
        'extBoundary.AddPoint(New Point3d(xyPt2.X, xyPt2.y, 0))
        'extBoundary.AddPoint(New Point3d(xyPt3.X, xyPt3.y, 0))
        'extBoundary.AddPoint(New Point3d(xyPt4.X, xyPt4.y, 0))
        extBoundary = acPoly.GeometricExtents
        Dim ptMin As Point3d = extBoundary.MinPoint
        Dim ptMax As Point3d = extBoundary.MaxPoint

        If xyGivenPoint.X >= ptMin.X And xyGivenPoint.X <= ptMax.X Then
            If xyGivenPoint.y >= ptMin.Y And xyGivenPoint.y <= ptMax.Y Then
                Return True
            End If
        End If
        Return False
    End Function
    Public Sub PR_DrawGloveDorsalZipper()
        Dim xyPALMER, xyDORSAL, xyDORSALWEB, xyStartEOS, xyEndEOS, xyTmp, xyZipStart, xyZipEnd, xyPt1, xyPt2, xyGivenPoint As ARMDIA1.XY
        Dim strProfile As String = ""
        Dim sSide As String = ""
        Dim nDirection As Integer
        Dim sAge As String = ""
        Dim sData As String = ""
        Dim nType, nAge, nInsertSize, nZipLength, nElasticProximal, nElastic As Double
        Dim EOStoSelectedPoint, EOStoCalculatedPoint, bIsLoop As Boolean
        Dim sZipLength As String = ""
        Dim sElasticProximal As String = ""
        Dim sWebOffSet As String = ""
        Dim nElasticFactor, nWebOffSet, nDrawnLength As Double

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim ptEntOpts As PromptEntityOptions = New PromptEntityOptions("Select Glove Profile")
        ptEntOpts.AllowNone = True
        ''ptEntOpts.Keywords.Add("Spline")
        Dim ptEntRes As PromptEntityResult = ed.GetEntity(ptEntOpts)
        If ptEntRes.Status <> PromptStatus.OK Then
            MsgBox("User Cancelled", 16, "Palm Zipper")
            Exit Sub
        End If

        Dim idEnt As ObjectId = ptEntRes.ObjectId
        Using tr As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acEnt As Entity = tr.GetObject(idEnt, OpenMode.ForRead)
            Dim rb As ResultBuffer = acEnt.GetXDataForApplication("ID")
            If Not IsNothing(rb) Then
                ' Get the values in the xdata
                For Each typeVal As TypedValue In rb
                    If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                        Continue For
                    End If
                    strProfile = strProfile & typeVal.Value
                Next
            End If
            If strProfile.Contains("Left") Then
                sSide = "Left"
                nDirection = 1
            ElseIf strProfile.Contains("Right") Then
                sSide = "Right"
                nDirection = -1
            Else
                MsgBox("A Glove Profile was not selected", 16, "Palm Zipper")
                Exit Sub
            End If

            ''// Get data for a DORSAL Zipper
            Dim nFound As Integer = 0
            Dim nReqired As Integer = 4
            Dim filterType(5) As TypedValue
            filterType.SetValue(New TypedValue(DxfCode.Operator, "<or"), 0)
            filterType.SetValue(New TypedValue(DxfCode.BlockName, "X Marker"), 1)
            filterType.SetValue(New TypedValue(DxfCode.Start, "Line"), 2)
            filterType.SetValue(New TypedValue(DxfCode.Operator, "or>"), 3)
            filterType.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "ID"), 4)
            filterType.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "Zipper"), 5)
            Dim selFilter As SelectionFilter = New SelectionFilter(filterType)
            Dim selResult As PromptSelectionResult = ed.SelectAll(selFilter)
            If selResult.Status <> PromptStatus.OK Then
                MsgBox("Can't get PALMER Zipper Details", 16, "Palm Zipper")
                Exit Sub
            End If
            Dim selectionSet As SelectionSet = selResult.Value
            For Each idObject As ObjectId In selectionSet.GetObjectIds()
                Dim acEntity As Entity = tr.GetObject(idObject, OpenMode.ForRead)
                Dim rbID As ResultBuffer = acEntity.GetXDataForApplication("ID")
                Dim strID As String = ""
                If Not IsNothing(rbID) Then
                    ' Get the values in the xdata
                    For Each typeVal As TypedValue In rbID
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        strID = strID & typeVal.Value
                    Next
                End If
                If strID.Equals(strProfile) = False Then
                    Continue For
                End If

                ''Dim ptPosition As Point3d = blkXMarker.Position
                Dim ptPosition As New Point3d
                If (acEntity.GetType().Name.ToUpper.Equals("BLOCKREFERENCE")) Then
                    Dim blkXMarker As BlockReference = TryCast(acEntity, BlockReference)
                    ptPosition = blkXMarker.Position
                End If
                Dim rbZipper As ResultBuffer = acEntity.GetXDataForApplication("Zipper")
                Dim strZipper As String = ""
                If Not IsNothing(rbZipper) Then
                    ' Get the values in the xdata
                    For Each typeVal As TypedValue In rbZipper
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        strZipper = strZipper & typeVal.Value
                    Next
                End If
                If strZipper.Equals("PALMER") Then
                    nFound = nFound + 1
                    ARMDIA1.PR_MakeXY(xyPALMER, ptPosition.X, ptPosition.Y)
                    Dim rbAge As ResultBuffer = acEntity.GetXDataForApplication("Age")
                    sAge = ""
                    If Not IsNothing(rbAge) Then
                        ' Get the values in the xdata
                        For Each typeVal As TypedValue In rbAge
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            sAge = sAge & typeVal.Value
                        Next
                    End If
                    Dim rbData As ResultBuffer = acEntity.GetXDataForApplication("Data")
                    sData = ""
                    If Not IsNothing(rbData) Then
                        ' Get the values in the xdata
                        For Each typeVal As TypedValue In rbData
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            sData = sData & typeVal.Value
                        Next
                    End If
                    nType = Val(Mid(sData, 10, 2))
                    nAge = Val(sAge)
                End If
                If strZipper.Equals("DORSAL") Then
                    nFound = nFound + 1
                    ARMDIA1.PR_MakeXY(xyDORSAL, ptPosition.X, ptPosition.Y)
                End If
                If strZipper.Equals("DORSAL-WEB") Then
                    nFound = nFound + 1
                    ARMDIA1.PR_MakeXY(xyDORSALWEB, ptPosition.X, ptPosition.Y)
                    Dim rbZipLen As ResultBuffer = acEntity.GetXDataForApplication("ZipperLength")
                    Dim sInsertSize As String = ""
                    If Not IsNothing(rbZipLen) Then
                        ' Get the values in the xdata
                        For Each typeVal As TypedValue In rbZipLen
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            sInsertSize = sInsertSize & typeVal.Value
                        Next
                    End If
                    nInsertSize = Val(sInsertSize)
                End If
                If strZipper.Equals("EOS") Then
                    nFound = nFound + 1
                    ''--------hEOS = hEnt 
                    ''-------GetGeometry(hEnt, & xyStartEOS, & xyEndEOS)
                    If (acEntity.GetType().Name.ToUpper.Equals("LINE")) Then
                        Dim acLine As Line = TryCast(acEntity, Line)
                        Dim ptStart As Point3d = acLine.StartPoint
                        Dim ptEnd As Point3d = acLine.EndPoint
                        ARMDIA1.PR_MakeXY(xyStartEOS, ptStart.X, ptStart.Y)
                        ARMDIA1.PR_MakeXY(xyEndEOS, ptEnd.X, ptEnd.Y)
                    End If
                End If
            Next

            ''// Check that sufficent data have been found, otherwise exit
            ''//
            If (nFound < nReqired) Then
                MsgBox("Missing data for selected Glove!", 16, "Palm Zipper")
                Exit Sub
            End If
            If (nFound > nReqired) Then
                MsgBox("Two or more drawings of the same style exist!" & Chr(10) & "Delete the extra drawing/s and try again.", 16, "Palm Zipper")
                Exit Sub
            End If

            ''// Create Dialog
            ''// Get Zipper style
            ''//
            Dim sUnits As String = "Inches"
            EOStoSelectedPoint = False
            EOStoCalculatedPoint = False

            ''// Proximal elastic
            '   sDlgElasticList =  "3/8\" Elastic\n3/4\" Elastic\n1½\" Elastic\nNo Elastic";
            '   if (nAge < 10)
            '   	sDlgElasticList =  "3/8\" Elastic\n" + sDlgElasticList;  // 1/2" for children under 10
            '   else
            '      	sDlgElasticList =  "3/4\" Elastic\n" + sDlgElasticList;  // Inch for adults 

            '// Length specification      	
            '   sDlgLengthList =  "Standard\nGive a length\nSelected Point";

            '// Offset from web
            '   sDlgWebOffSetList =  "1-1/8\"\n3/4\"";
            '   sDlgWebOffSetList =  "1-1/8\"\n"  + sDlgWebOffSetList ;	// Set Default
            bIsLoop = True
            While (bIsLoop)
                '               nButX = 65; nButY = 55;
                '           hDlg = Open("dialog", sSide + " Glove Zipper (DORSAL)", "font Helv 8", 20, 20, 210, 75);

                'AddControl(hDlg, "pushbutton", nButX, nButY, 35, 14, "Cancel", "%cancel", "");
                'AddControl(hDlg, "pushbutton", nButX + 48, nButY, 35, 14, "OK", "%ok", "");

                'AddControl(hDlg, "ltext", 5, 12, 28, 14, "EOS to", "string", "");
                'AddControl(hDlg, "combobox", 35, 10, 70, 40, sDlgLengthList, "string", "sZipLength");

                'AddControl(hDlg, "ltext", 110, 12, 30, 14, "Proximal:", "string", "");
                '	AddControl(hDlg, "dropdownlist", 145, 10, 60, 70, sDlgElasticList, "string", "sElasticProximal");

                'AddControl(hDlg, "ltext", 100, 32, 50, 14, "Below Web:", "string", "");
                '	AddControl(hDlg, "dropdownlist", 145, 30, 60, 40, sDlgWebOffSetList, "string", "sWebOffSet");

                '     	Ok = Display("dialog", hDlg, "%center");
                ' 	Close("dialog", hDlg);

                'If (Ok == %cancel ) Exit (%ok, "User Cancel!") ;
                Dim oZipperPalmFrm As New ZIPPALM_frm
                oZipperPalmFrm.g_sCaption = sSide + " Glove Zipper (DORSAL)"
                oZipperPalmFrm.g_nAge = nAge
                oZipperPalmFrm.ShowDialog()
                If oZipperPalmFrm.g_bIsCancel = True Then
                    MsgBox("User Cancel!", 16, "Palm Zipper")
                    Exit Sub
                End If
                sZipLength = oZipperPalmFrm.g_sZipLength
                sElasticProximal = oZipperPalmFrm.g_sElasticProximal
                sWebOffSet = oZipperPalmFrm.g_sWebOffset

                If sZipLength.Equals("Selected Point") Then
                    EOStoSelectedPoint = True
                End If
                If sZipLength.Equals("Standard") Then
                    EOStoCalculatedPoint = True
                End If
                If (EOStoSelectedPoint = True Or EOStoCalculatedPoint = True) Then
                    bIsLoop = False
                Else
                    nZipLength = Val(sZipLength)
                    If (nZipLength = 0 And sZipLength.Length > 0) Then
                        MsgBox("Invalid given length!" & Chr(10) & Chr(10) & "To use this option, type over the text in the " &
                               Chr(34) & "EOS to: " & Chr(34) & "box with the required length in Inches (Decimal Inches).", 16, "Palm Zipper")
                        bIsLoop = True
                    Else
                        bIsLoop = False
                    End If
                End If
            End While
            ''// Establish allowance for zippers
            ''//
            nElasticProximal = 0.75
            If (String.Compare(Mid(sElasticProximal, 1, 1), "N") = 0) Then
                nElasticProximal = 0
            End If
            If (String.Compare(Mid(sElasticProximal, 1, 3), "3/8") = 0) Then
                nElasticProximal = 0.375
            End If
            If (String.Compare(Mid(sElasticProximal, 1, 1), "1") = 0) Then
                nElasticProximal = 1.5
            End If
            nElastic = nElasticProximal
            ''// Elastic factor
            ''// Set by age
            If (nAge <= 6) Then
                nElasticFactor = 1
            Else
                nElasticFactor = 0.95     ''// Glove to elbow & Normal Glove
                If (nType = 2) Then
                    nElasticFactor = 0.92   ''// Glove to axilla
                End If
            End If

            ''// Establish Minimum offset from web    
            nWebOffSet = 1.125
            If (String.Compare(Mid(sWebOffSet, 1, 1), "3") = 0) Then
                nWebOffSet = 0.75
            End If
        End Using

        ''// Draw on layer Note
        ARMDIA1.PR_SetLayer("Notes")
        ''// Draw Zipper
        ''// 
        ''// Establish EOS start And end
        ''//
        Dim nLengthToStart As Double = ARMDIA1.FN_CalcLength(xyStartEOS, xyPALMER)
        Dim nLengthToEnd As Double = ARMDIA1.FN_CalcLength(xyEndEOS, xyPALMER)
        If (nLengthToStart > nLengthToEnd) Then
            ''// Swap start and end
            xyTmp = xyStartEOS
            xyStartEOS = xyEndEOS
            xyEndEOS = xyTmp
        End If

        ''// Angles
        ''//
        Dim aEOS As Double = ARMDIA1.FN_CalcAngle(xyStartEOS, xyEndEOS)
        ''// Need to offset xyDORSAL by 1/8th as the point lies on the thumb arc
        ''//
        ''xyDORSAL = CalcXY ("relpolar", xyDORSAL, 0.125, aEOS -180) ;    
        ARMDIA1.PR_CalcPolar(xyDORSAL, aEOS - 180, 0.125, xyDORSAL)
        Dim aDORSAL As Double = ARMDIA1.FN_CalcAngle(xyStartEOS, xyDORSAL)
        ''// Zipper start point on EOS
        ''// 
        Dim aAngle As Double = aEOS - aDORSAL
        Dim nLength As Double = ARMDIA1.FN_CalcLength(xyStartEOS, xyDORSAL)
        Dim nB As Double = System.Math.Abs(System.Math.Cos(aAngle) * nLength)
        ''xyZipStart = CalcXY("relpolar", xyStartEOS, nB, aEOS)
        ''ARMDIA1.PR_CalcPolar(xyStartEOS, aEOS, nB / 2, xyZipStart)
        ARMDIA1.PR_CalcPolar(xyStartEOS, aEOS, (nB + 0.625) / 4, xyZipStart)

        Dim aDORSALWEB As Double = ARMDIA1.FN_CalcAngle(xyStartEOS, xyDORSALWEB)
        aAngle = System.Math.Abs(aEOS - aDORSALWEB)
        nLength = ARMDIA1.FN_CalcLength(xyStartEOS, xyDORSALWEB)
        Dim nZipConstructLen As Double = System.Math.Abs(System.Math.Sin(aAngle) * nLength)

        Dim aZipper As Double = aEOS + (90 * nDirection)
        If (EOStoCalculatedPoint) Then
            ''xyZipEnd = CalcXY ("relpolar", xyZipStart, nZipConstructLen  - nWebOffSet - nInsertSize , aZipper) 
            ARMDIA1.PR_CalcPolar(xyZipStart, aZipper, nZipConstructLen - nWebOffSet - nInsertSize, xyZipEnd)
        End If
        If (EOStoSelectedPoint) Then
            ''// Only allow the point if it lies within the rectangle formed by the
            ''	// EOS and the xyDORSALWEB plus a 2" tolerance
            ''// EOS = length and angle, xyDORSALWEB = width

            ''xyPt1 = CalcXY ("relpolar", xyStartEOS, nZipConstructLen + .0625  , aZipper) ; 
            ARMDIA1.PR_CalcPolar(xyStartEOS, aZipper, nZipConstructLen + 0.0625, xyPt1)
            ''xyPt1 = CalcXY("relpolar", xyPt1, 2, (aEOS + (nDirection * 180))) ; 
            ARMDIA1.PR_CalcPolar(xyPt1, (aEOS + (nDirection * 180)), 2, xyPt1)
            ''xyPt2 = CalcXY("relpolar", xyEndEOS, nZipConstructLen + 0.0625, aZipper) ; 
            ARMDIA1.PR_CalcPolar(xyEndEOS, aZipper, nZipConstructLen + 0.0625, xyPt2)
            ''xyPt2 = CalcXY("relpolar", xyPt2, 2, aEOS)
            ARMDIA1.PR_CalcPolar(xyPt2, aEOS, 2, xyPt2)

            bIsLoop = True
            While (bIsLoop)
                Dim pPtRes As PromptPointResult
                Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                pPtOpts.Message = vbLf & "Select End of Zipper "
                pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                If pPtRes.Status = PromptStatus.Cancel Then
                    MsgBox("End Not Selected", 16, "Palm Zipper")
                    Exit Sub
                End If
                Dim ptEnd As Point3d = pPtRes.Value
                ARMDIA1.PR_MakeXY(xyGivenPoint, ptEnd.X, ptEnd.Y)
                ''If (!Calc("inpoly", xyGivenPoint, xyStartEOS, xyPt1, xyPt2, xyEndEOS)) Then
                If (PR_CheckPointExistInside(xyGivenPoint, xyStartEOS, xyPt1, xyPt2, xyEndEOS) = False) Then
                    If MsgBox("Given point can't be used! Try again." & Chr(10) & "Or use cancel to Exit", 1, "Palm Zipper") = MsgBoxResult.Cancel Then
                        MsgBox("User Cancelled", 16, "Palm Zipper")
                        Exit Sub
                    End If
                Else
                    bIsLoop = False
                End If
            End While
            aDORSALWEB = ARMDIA1.FN_CalcAngle(xyStartEOS, xyGivenPoint)
            aAngle = System.Math.Abs(aEOS - aDORSALWEB)
            nLength = ARMDIA1.FN_CalcLength(xyStartEOS, xyGivenPoint)
            Dim nA As Double = System.Math.Abs(System.Math.Sin(aAngle) * nLength)
            ''xyZipEnd = CalcXY("relpolar", xyZipStart, nA, aZipper)
            ARMDIA1.PR_CalcPolar(xyZipStart, aZipper, nA, xyZipEnd)
            ''// Ensure that selected point Is no closer than nWebOffSet to the web
            ''// within 1/8th of an inch
            nLength = ARMDIA1.FN_CalcLength(xyZipStart, xyZipEnd) - nZipConstructLen
            If ((System.Math.Abs(nLength) < (nWebOffSet - 0.125)) Or (nLength > 0)) Then
                Dim sText As String = ""
                If (nLength > 0) Then
                    sText = "The end of the zipper will above the nearest web\\slant insert."
                Else
                    sText = "The end of the zipper will be closer to the nearest web\\slant insert than " +
                        Format("length", nWebOffSet) + ".  Actual distance is " + Format("length", System.Math.Abs(nLength))
                End If
                Dim ok As MsgBoxResult = MsgBox(sText & Chr(10) & "Use YES to use this point or" & Chr(10) & "Use NO to default to Standard.", 3, "Palm Zipper")
                If ok = MsgBoxResult.Cancel Then
                    MsgBox("User Cancelled", 16, "Palm Zipper")
                End If
                If ok = MsgBoxResult.No Then
                    ''xyZipEnd = CalcXY ("relpolar", xyZipStart, nZipConstructLen  - nWebOffSet, aZipper)
                    ARMDIA1.PR_CalcPolar(xyZipStart, aZipper, nZipConstructLen - nWebOffSet, xyZipEnd)
                End If
            End If
        End If
        If (nZipLength > 0) Then
            ''xyZipEnd = CalcXY ("relpolar", xyZipStart, (nZipLength * nElasticFactor) - nElastic, aZipper)
            ARMDIA1.PR_CalcPolar(xyZipStart, aZipper, (nZipLength * nElasticFactor) - nElastic, xyZipEnd)
            nDrawnLength = ARMDIA1.FN_CalcLength(xyZipStart, xyZipEnd)
            nLength = nDrawnLength - nZipConstructLen
            If ((System.Math.Abs(nLength) < (nWebOffSet - 0.125)) Or (nLength > 0)) Then
                Dim status As MsgBoxResult = MsgBox("The end of the zipper will be above or closer to the nearest web\\slant insert than" +
                                                    Format("length", nWebOffSet) + " with the given length." +
                                                    Chr(10) + "Use OK to default to Standard.", 1, "Palm Zipper")
                If status = MsgBoxResult.Cancel Then
                    MsgBox("User Cancelled", 16, "Palm Zipper")
                End If
                ''xyZipEnd = CalcXY ("relpolar", xyZipStart, nZipConstructLen  - nWebOffSet, aZipper) 
                ARMDIA1.PR_CalcPolar(xyZipStart, aZipper, nZipConstructLen - nWebOffSet, xyZipEnd)
                nDrawnLength = ARMDIA1.FN_CalcLength(xyZipStart, xyZipEnd)
                nZipLength = (nDrawnLength + nElastic) / nElasticFactor
            End If
        Else
            nDrawnLength = ARMDIA1.FN_CalcLength(xyZipStart, xyZipEnd)
            nZipLength = (nDrawnLength + nElastic) / nElasticFactor
        End If
        sZipLength = Format("length", nZipLength)

        ''// Draw markers and Text label
        ''// Each entity Is given an ID Data Base value that links them together
        ''// This Is based on the UID of the symbol "ZipperText" And the Glove Profile ID.
        ''// A Symbol Is used as DB values can't be attached to text entities 

        ''// Add text symbol
        ''xyTmp = CalcXY("relpolar", xyZipStart, nDrawnLength / 2, aZipper)
        ARMDIA1.PR_CalcPolar(xyZipStart, aZipper, nDrawnLength / 2, xyTmp)
        ''// Create ID string
        Dim sZipperID As String = strProfile ''+ MakeString("scalar", UID("get", hEnt)) ;
        PR_DrawTextAsSymbol(xyTmp, sZipperID, sZipLength + " Dorsal", aZipper - 180)
        '// Label entity with ID string and make Zipper %true
        SetDBData("ID", sZipperID)
        SetDBData("Zipper", "1")
        SetDBData("Data", sZipLength + " Dorsal")

        ''// Add label and arrows
        ''hEnt = AddEntity("marker", "closed arrow", xyZipStart, 0.5, 0.125, aEOS + (90 * nDirection))
        PR_DrawClosedArrow(xyZipStart, aEOS + (90 * nDirection))
        SetDBData("ID", sZipperID)
        SetDBData("Zipper", "1")

        ''hEnt = AddEntity("marker","closed arrow", xyZipEnd , 0.5 ,0.125, aEOS - (90 * nDirection)) ;	
        PR_DrawClosedArrow(xyZipEnd, aEOS - (90 * nDirection))
        SetDBData("ID", sZipperID)
        SetDBData("Zipper", "1")

        PR_DrawLine(xyZipStart, xyZipEnd)
        SetDBData("ID", sZipperID)
        SetDBData("Zipper", "1")

        ''// Reset And exit
        ''Execute("menu", "SetLayer", Table("find", "layer", "1")) 
        MsgBox("Zipper drawing Complete", 0, "Dorsal Zipper")
    End Sub
    Public Sub PR_DrawWaistChapPantyZipper()
        ''---------WHZCPANT.D-----------------
        Dim strProfile As String = ""
        Dim sStyleID As String = ""
        Dim sLeg As String = ""
        Dim xyOtemplate, xyFold, xyAnkle, xyEOS, xySmallestCir, xyTmp, xyHorizStart, xyHorizEnd As ARMDIA1.XY
        Dim sUnits As String = ""
        Dim sTemplate As String = ""
        Dim sType As String = ""
        Dim sElasticNote As String = ""
        Dim EOS, bIsLoop, bIsSelect As Boolean
        Dim nMedial As Integer = 0
        Dim sZipLength As String = ""
        Dim sElasticProximal As String = ""
        Dim sElasticDistal As String = ""
        Dim nZipLength, nElasticProximal, nElasticDistal, nElastic, nZipOff, nZipperInChap, nDrawnLength As Double
        Dim xyZipperStartInChap, xyZipperEndInChap As ARMDIA1.XY

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim ptEntOpts As PromptEntityOptions = New PromptEntityOptions(Chr(10) & "Select a Leg Profile")
        ptEntOpts.AllowNone = True
        ''ptEntOpts.Keywords.Add("Spline")
        ptEntOpts.SetRejectMessage(Chr(10) & "Only Curve")
        ptEntOpts.AddAllowedClass(GetType(Spline), True)
        Dim ptEntRes As PromptEntityResult = ed.GetEntity(ptEntOpts)
        If ptEntRes.Status <> PromptStatus.OK Then
            MsgBox("User Cancelled", 16, "Knee Contracture")
            Exit Sub
        End If
        Dim idEnt As ObjectId = ptEntRes.ObjectId
        Dim ptCurveColl As New Point3dCollection
        Dim fitPtsCount As Double = 0
        Using tr As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acSpline As Spline = tr.GetObject(idEnt, OpenMode.ForRead)
            Dim rb As ResultBuffer = acSpline.GetXDataForApplication("ID")
            If Not IsNothing(rb) Then
                ' Get the values in the xdata
                For Each typeVal As TypedValue In rb
                    If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                        Continue For
                    End If
                    strProfile = strProfile & typeVal.Value
                Next
            End If
            ''// Check if this is a valid Leg Curve exit if not
            strProfile = Trim(strProfile)
            If strProfile.Contains("LeftLegCurve") Then
                sStyleID = Mid(strProfile, 1, strProfile.Length - 12)
                sLeg = "Left"
            ElseIf strProfile.Contains("RightLegCurve") Then
                sStyleID = Mid(strProfile, 1, strProfile.Length - 13)
                sLeg = "Right"
            Else
                MsgBox("A Leg Profile was not selected" & Chr(10), 16, "Ankle Zipper")
                Exit Sub
            End If
            fitPtsCount = acSpline.NumFitPoints
            Dim i As Integer = 0
            While (i < fitPtsCount)
                ptCurveColl.Add(acSpline.GetFitPointAt(i))
                i = i + 1
            End While

            ''// Get Marker data
            sStyleID = sStyleID + sLeg
            Dim sOtemplateID As String = sStyleID + "Origin"
            Dim sFoldID As String = sStyleID + "Fold"
            Dim sAnkleID As String = sStyleID + "Ankle"
            Dim sEOSID As String = sStyleID + "EOS"
            Dim sNoteID As String = sStyleID + "PantyElasticNote"

            'Dim filterType(10) As TypedValue
            'filterType.SetValue(New TypedValue(DxfCode.Operator, "<or"), 0)
            'filterType.SetValue(New TypedValue(DxfCode.Operator, "<and"), 1)
            'filterType.SetValue(New TypedValue(DxfCode.BlockName, "TEXTASSYMBOL"), 2)
            'filterType.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "ID"), 3)
            'filterType.SetValue(New TypedValue(DxfCode.Operator, "and>"), 4)
            'filterType.SetValue(New TypedValue(DxfCode.BlockName, "X Marker"), 5)
            'filterType.SetValue(New TypedValue(DxfCode.Operator, "<or"), 6)
            'filterType.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "OriginID"), 7)
            'filterType.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "ID"), 8)
            'filterType.SetValue(New TypedValue(DxfCode.Operator, "or>"), 9)
            'filterType.SetValue(New TypedValue(DxfCode.Operator, "or>"), 10)

            Dim filterType(3) As TypedValue
            filterType.SetValue(New TypedValue(DxfCode.Operator, "<or"), 0)
            filterType.SetValue(New TypedValue(DxfCode.BlockName, "TEXTASSYMBOL"), 1)
            filterType.SetValue(New TypedValue(DxfCode.BlockName, "X Marker"), 2)
            filterType.SetValue(New TypedValue(DxfCode.Operator, "or>"), 3)
            Dim selFilter As SelectionFilter = New SelectionFilter(filterType)
            Dim selResult As PromptSelectionResult = ed.SelectAll(selFilter)
            If selResult.Status <> PromptStatus.OK Then
                MsgBox("Can't find Marker Details", 16, "Ankle Zipper")
                Exit Sub
            End If
            Dim selectionSet As SelectionSet = selResult.Value
            Dim nMarkersFound As Double = 0
            Dim nMarkersRequired As Double = 3
            Dim Panty As Boolean = True
            For Each idObject As ObjectId In selectionSet.GetObjectIds()
                Dim blkXMarker As BlockReference = tr.GetObject(idObject, OpenMode.ForRead)
                Dim position As Point3d = blkXMarker.Position
                If position.X = 0 And position.Y = 0 Then
                    Continue For
                End If
                Dim resbuf As ResultBuffer = blkXMarker.GetXDataForApplication("OriginID")
                Dim strXMarkerData As String = ""
                If Not IsNothing(resbuf) Then
                    For Each typeVal As TypedValue In resbuf
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        strXMarkerData = strXMarkerData & typeVal.Value
                    Next
                End If
                If strXMarkerData.Equals(sOtemplateID) Then
                    nMarkersFound = nMarkersFound + 1
                    ARMDIA1.PR_MakeXY(xyOtemplate, position.X, position.Y)
                    Dim resbufUnits As ResultBuffer = blkXMarker.GetXDataForApplication("units")
                    sUnits = ""
                    If Not IsNothing(resbufUnits) Then
                        For Each typeVal As TypedValue In resbufUnits
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            sUnits = sUnits & typeVal.Value
                        Next
                    End If
                    Dim resbufPressure As ResultBuffer = blkXMarker.GetXDataForApplication("Pressure")
                    sTemplate = ""
                    If Not IsNothing(resbufPressure) Then
                        For Each typeVal As TypedValue In resbufPressure
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            sTemplate = sTemplate & typeVal.Value
                        Next
                    End If
                    Dim resbufData As ResultBuffer = blkXMarker.GetXDataForApplication("Data")
                    sType = ""
                    If Not IsNothing(resbufData) Then
                        For Each typeVal As TypedValue In resbufData
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            sType = sType & typeVal.Value
                        Next
                    End If
                End If
                Dim rbID As ResultBuffer = blkXMarker.GetXDataForApplication("ID")
                Dim strID As String = ""
                If Not IsNothing(rbID) Then
                    For Each typeVal As TypedValue In rbID
                        If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                            Continue For
                        End If
                        strID = strID & typeVal.Value
                    Next
                End If
                If strID.Equals(sFoldID) Then
                    nMarkersFound = nMarkersFound + 1
                    ARMDIA1.PR_MakeXY(xyFold, position.X, position.Y)
                End If
                If strID.Equals(sNoteID) Then
                    Dim rbData As ResultBuffer = blkXMarker.GetXDataForApplication("Data")
                    sElasticNote = ""
                    If Not IsNothing(rbData) Then
                        For Each typeVal As TypedValue In rbData
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            sElasticNote = sElasticNote & typeVal.Value
                        Next
                    End If
                End If
                If strID.Equals(sAnkleID) Then
                    ARMDIA1.PR_MakeXY(xyAnkle, position.X, position.Y)
                    Panty = False
                End If
                If strID.Equals(sEOSID) Then
                    nMarkersFound = nMarkersFound + 1
                    ARMDIA1.PR_MakeXY(xyEOS, position.X, position.Y)
                End If
            Next
            ''// Check if that the markers have been found, otherwise exit
            If sType.Contains("CHAP") = False Then
                MsgBox("Can't use this zipper for a Waist Ht style!", 16, "Waist Chap Ankle Zipper")
                Exit Sub
            End If
            If (nMarkersFound < nMarkersRequired) Then
                MsgBox("Missing markers for selected leg profile, data not found!", 16, "Waist Chap Ankle Zipper")
                Exit Sub
            End If
            If (nMarkersFound > nMarkersRequired) Then
                MsgBox("Two or more drawings of the same style exist!" & Chr(10) & "Delete the extra drawing/s and try again.", 16, "Ankle Zipper")
                Exit Sub
            End If
            If Panty = False Then
                MsgBox("This style has a foot so it is not a Panty!." & Chr(10) & "Delete the extra drawing/s and try again.", 16, "Ankle Zipper")
                Exit Sub
            End If

            ''// Get Smallest circumference by looping through the vetex of the curve 
            ''//
            xySmallestCir.y = 10000  ''// Impossibly large value used As an initial value  
            Dim nn As Integer = 0
            While (nn < ptCurveColl.Count)
                ARMDIA1.PR_MakeXY(xyTmp, ptCurveColl(nn).X, ptCurveColl(nn).Y)
                If (xyTmp.y < xySmallestCir.y) Then
                    xySmallestCir.y = xyTmp.y
                Else
                    Exit While
                End If
                nn = nn + 1
            End While

            ''// Create Dialog
            ''// Get Zipper style
            ''//
            sUnits = "Inches"
            EOS = False
            bIsLoop = True
            bIsSelect = False
            '         sDlgElasticList = "3/8\" Elastic\n3/4\" Elastic\n1½\" Elastic\nNo Elastic";
            'sDlgProximalElasticList =  "No Elastic\n" + sDlgElasticList;
            'if (StringCompare(sElasticNote, "NO ELASTIC"))
            '	sDlgDistalElasticList =  "No Elastic\n" + sDlgElasticList;
            'else
            '	sDlgDistalElasticList =  "3/4\" Elastic\n" + sDlgElasticList;
            'sDlgLengthList = "EOS\nGive a length\nSelected Point\n";
            nMedial = 0
            While (bIsLoop)
                'nButX = 70; nButY = 65;
                'hDlg = Open("dialog", sLeg + " Distal EOS Zipper (CHAP)", "font Helv 8", 20, 20, 235, 95);

                'AddControl(hDlg, "pushbutton", nButX, nButY, 35, 14, "Cancel", "%cancel", "");
                'AddControl(hDlg, "pushbutton", nButX + 48, nButY, 35, 14, "OK", "%ok", "");

                'AddControl(hDlg, "ltext", 5, 12, 55, 14, "Distal EOS to", "string", "");
                'AddControl(hDlg, "combobox", 60, 10, 70, 40, sDlgLengthList, "string", "sZipLength");

                'AddControl(hDlg, "ltext", 135, 12, 25, 14, "Distal :", "string", "");
                'AddControl(hDlg, "dropdownlist", 170, 10, 60, 70, sDlgDistalElasticList, "string", "sElasticDistal");

                'AddControl(hDlg, "ltext", 135, 30, 30, 14, "Proximal:", "string", "");
                '	AddControl(hDlg, "dropdownlist", 170, 28, 60, 70, sDlgProximalElasticList, "string", "sElasticProximal");

                ' 	AddControl(hDlg, "checkbox", nButX + 12, nButY - 18, 65, 15, "Medial Zipper", "number", "nMedial");

                '  	Ok = Display("dialog", hDlg, "%center");
                ' 	Close("dialog", hDlg);

                'If (Ok == %cancel ) Exit (%ok, "User Cancel!") ;

                Dim oWHChapPantZipFrm As New WHZCPANT_frm
                oWHChapPantZipFrm.g_sCaption = sLeg + " Distal EOS Zipper (CHAP)"
                oWHChapPantZipFrm.g_sElasticNote = sElasticNote
                oWHChapPantZipFrm.ShowDialog()
                If oWHChapPantZipFrm.g_bIsCancel = True Then
                    MsgBox("User Cancel!", 16, "Distal EOS Zipper")
                    Exit Sub
                End If
                sZipLength = oWHChapPantZipFrm.g_sZipLength
                sElasticProximal = oWHChapPantZipFrm.g_sElasticProximal
                sElasticDistal = oWHChapPantZipFrm.g_sElasticDistal
                If oWHChapPantZipFrm.g_bIsMedialZipper = True Then
                    nMedial = 1
                End If

                If sZipLength.Equals("EOS") = True Or sZipLength.Equals("Selected Point") = True Then
                    bIsLoop = False
                    If sZipLength.Equals("EOS") = True Then
                        EOS = True
                    Else
                        bIsSelect = True
                    End If
                Else
                    nZipLength = Val(sZipLength)
                    If (nZipLength = 0 And sZipLength.Length > 0) Then
                        bIsLoop = True
                    Else
                        bIsLoop = False
                        EOS = False
                        bIsSelect = False
                    End If
                End If
            End While

            If (bIsSelect Or nZipLength > 0) Then ''//closed zipper at proximal end
                nElasticProximal = 0
            Else
                nElasticProximal = 0.75
                If (String.Compare(Mid(sElasticProximal, 1, 1), "N") = 0) Then
                    nElasticProximal = 0
                End If
                If (String.Compare(Mid(sElasticProximal, 1, 3), "3/8") = 0) Then
                    nElasticProximal = 0.375
                End If
                If (String.Compare(Mid(sElasticProximal, 1, 1), "1") = 0) Then
                    nElasticProximal = 1.5
                End If
            End If

            nElasticDistal = 0.75
            If (String.Compare(Mid(sElasticDistal, 1, 1), "N") = 0) Then
                nElasticDistal = 0
            End If
            If (String.Compare(Mid(sElasticDistal, 1, 3), "3/8") = 0) Then
                nElasticDistal = 0.375
            End If
            If (String.Compare(Mid(sElasticDistal, 1, 1), "1") = 0) Then
                nElasticDistal = 1.5
            End If
            nElastic = nElasticDistal + nElasticProximal

            Dim nInchToCM As Double = 1
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                nInchToCM = 2.54
            End If
            nZipOff = 1.125 * nInchToCM
            ARMDIA1.PR_MakeXY(xyHorizStart, ptCurveColl(0).X, ptCurveColl(0).Y)
            xyHorizStart.y = xySmallestCir.y - nZipOff
            xyHorizEnd.y = xyHorizStart.y

            ''//if  EOS given
            If (EOS = True) Then
                If (nMedial = 1) Then
                    xyHorizEnd.X = xyFold.X
                    nElastic = nElasticDistal
                    nZipperInChap = 0
                Else
                    xyZipperStartInChap.X = xyFold.X
                    xyZipperStartInChap.y = xyFold.y - 1.125
                    xyZipperEndInChap = xyEOS
                    xyZipperEndInChap.y = xyFold.y - 1.125
                    nZipperInChap = ARMDIA1.FN_CalcLength(xyZipperStartInChap, xyZipperEndInChap)
                    xyHorizEnd.X = xyFold.X
                End If
                xyHorizEnd.y = xyHorizStart.y
            Else
                ''// No elastic at proximal as this implies a closed zipper
                nElastic = nElasticDistal
                ''//create a dummy xyHorizEnd for use on insertion
                If (xyHorizStart.X + (nZipLength / 1.2) > xyFold.X) Then
                    xyHorizEnd.X = xyFold.X
                Else
                    xyHorizEnd.X = xyHorizStart.X + (nZipLength / 1.2)
                End If
                xyHorizEnd.y = xyHorizStart.y
            End If

            ''// For Medial Zipper step back 3 tape
            If (nMedial = 1) Then
                Dim nFoldOffset As Double = 0
                Dim nTemplate As Double = Val(Mid(sTemplate, 1, 2))
                If (nTemplate >= 30) Then
                    nFoldOffset = 1.25 * 3
                End If
                If (nTemplate = 13) Then
                    nFoldOffset = 1.31 * 3
                    MsgBox("WARNING!" & Chr(10) & "Template drawn using 13 Down Stretch!", 16, "Waist Distal Chap Zipper")
                End If
                If (nTemplate = 9) Then
                    nFoldOffset = 1.37 * 3
                End If
                xyHorizEnd.X = xyHorizEnd.X - nFoldOffset
            End If

            ''// If select given then prompt user for end of zip
            ''//
            If (bIsSelect) Then
                Dim pEndPtRes As PromptPointResult
                Dim pEndPtOpts As PromptPointOptions = New PromptPointOptions("")
                pEndPtOpts.Message = vbLf & "Select End of Zipper"
                pEndPtRes = acDoc.Editor.GetPoint(pEndPtOpts)
                If pEndPtRes.Status = PromptStatus.Cancel Then
                    MsgBox("End not selected", 16, "Ankle Zipper")
                    Exit Sub
                End If
                Dim ptEndZip As Point3d = pEndPtRes.Value
                ARMDIA1.PR_MakeXY(xyHorizEnd, ptEndZip.X, ptEndZip.Y)
                xyHorizEnd.y = xyHorizStart.y
                If (xyHorizEnd.X > xyFold.X) Then
                    xyZipperStartInChap.X = xyFold.X
                    xyZipperStartInChap.y = xyFold.y - 1.125
                    xyZipperEndInChap.X = xyHorizEnd.X
                    xyZipperEndInChap.y = xyFold.y - 1.125
                    If (xyZipperEndInChap.X < xyFold.X) Then
                        nZipperInChap = 0
                    Else
                        nZipperInChap = ARMDIA1.FN_CalcLength(xyZipperStartInChap, xyZipperEndInChap)
                    End If
                    xyHorizEnd.X = xyFold.X
                    nElastic = 0
                    EOS = True
                End If
            End If
        End Using
        ''// Revise radius to include Zip offset
        ''//
        ARMDIA1.PR_SetLayer("Notes")
        Dim xyBlock As ARMDIA1.XY
        ARMDIA1.PR_MakeXY(xyBlock, xyHorizStart.X + (xyHorizEnd.X - xyHorizStart.X) / 2, xyHorizStart.y + 0.125)

        ''hTextEnt = AddEntity("symbol", "TextAsSymbol", xyHorizStart.X + (xyHorizEnd.X - xyHorizStart.X) / 2, xyHorizStart.y + 0.125, 1, 1, 0) ;  

        ''// Create ID string
        Dim sZipperID As String = sStyleID ''+ MakeString("scalar", UID("get", hTextEnt)) ;
        '     SetDBData(hTextEnt, "ID", sZipperID);
        'SetDBData(hTextEnt, "Zipper", "1");
        'SetDBData(hTextEnt, "Data", "INVALID ZIPPER CALCULATION"); //Do this just incase it fails
        If (nZipLength > 0) Then
            xyHorizEnd.X = xyHorizStart.X + ((nZipLength / 1.2))
            xyHorizEnd.y = xyHorizStart.y
        Else
            nDrawnLength = ARMDIA1.FN_CalcLength(xyHorizStart, xyHorizEnd)
            nZipLength = (nDrawnLength + nElastic) * 1.2
        End If
        PR_DrawLine(xyHorizStart, xyHorizEnd)
        SetDBData("ID", sZipperID)
        SetDBData("Zipper", "1")
        PR_DrawClosedArrow(xyHorizStart, 0)
        SetDBData("ID", sZipperID)
        SetDBData("Zipper", "1")

        ''// Add label and arrows
        ''// If a zipper extends to the EOS then we draw this bit
        If (nZipperInChap <> 0) Then
            PR_DrawLine(xyZipperStartInChap, xyZipperEndInChap)
            SetDBData("ID", sZipperID)
            SetDBData("Zipper", "1")
            PR_DrawClosedArrow(xyZipperStartInChap, 0)
            SetDBData("ID", sZipperID)
            SetDBData("Zipper", "1")
            PR_DrawClosedArrow(xyZipperEndInChap, 180)
            SetDBData("ID", sZipperID)
            SetDBData("Zipper", "1")

            ''// Get Velcro overlap (If it exists!) and delete it.
            ''-----------sTmp = "DB ID ='" + sStyleID + "' AND DB Zipper ='VelcroOverlap'";
            ''-----------hChan = Open("selection", sTmp)
            Dim filterVelcro(1) As TypedValue
            filterVelcro.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "ID"), 0)
            filterVelcro.SetValue(New TypedValue(DxfCode.ExtendedDataRegAppName, "Zipper"), 1)
            Dim selVelcroFilter As SelectionFilter = New SelectionFilter(filterVelcro)
            Dim selVelcroResult As PromptSelectionResult = ed.SelectAll(selVelcroFilter)
            If selVelcroResult.Status <> PromptStatus.OK Then
                MsgBox("Can't find Marker Details", 16, "Ankle Zipper")
                Exit Sub
            End If
            Dim selectionSetVelcro As SelectionSet = selVelcroResult.Value
            Using acTr As Transaction = acCurDb.TransactionManager.StartTransaction()
                For Each idObject As ObjectId In selectionSetVelcro.GetObjectIds()
                    Dim acVelcroEnt As Entity = acTr.GetObject(idObject, OpenMode.ForWrite)
                    Dim rbID As ResultBuffer = acVelcroEnt.GetXDataForApplication("ID")
                    Dim strID As String = ""
                    If Not IsNothing(rbID) Then
                        ' Get the values in the xdata
                        For Each typeVal As TypedValue In rbID
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            strID = strID & typeVal.Value
                        Next
                    End If
                    If strID.Contains(sStyleID) = False Then
                        Continue For
                    End If

                    Dim rbZipper As ResultBuffer = acVelcroEnt.GetXDataForApplication("Zipper")
                    Dim strZipper As String = ""
                    If Not IsNothing(rbZipper) Then
                        ' Get the values in the xdata
                        For Each typeVal As TypedValue In rbZipper
                            If typeVal.TypeCode = DxfCode.ExtendedDataRegAppName Then
                                Continue For
                            End If
                            strZipper = strZipper & typeVal.Value
                        Next
                    End If
                    If strZipper.Contains("VelcroOverlap") = False Then
                        Continue For
                    End If
                    Dim sLayer As String = acVelcroEnt.Layer
                    If (sLayer.ToUpper().Equals("NOTES") And acVelcroEnt.GetType().Name.ToUpper.Equals("LINE")) Then
                        Dim acLyrTbl As LayerTable
                        acLyrTbl = acTr.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                        If acLyrTbl.Has("Template" + sLeg) = False Then
                            Continue For
                        End If
                        acVelcroEnt.SetLayerId(acLyrTbl("Template" + sLeg), True)
                        acVelcroEnt.XData = New ResultBuffer(New TypedValue(DxfCode.ExtendedDataRegAppName, "Zipper"))
                    Else
                        acVelcroEnt.Erase()
                    End If
                Next
                acTr.Commit()
            End Using
        Else
            PR_DrawClosedArrow(xyHorizEnd, 180)
            SetDBData("ID", sZipperID)
            SetDBData("Zipper", "1")
        End If

        ''// Update text symbol with Zipper Length
        Dim sZipperText As String = ""
        If (nMedial = 1) Then
            sZipperText = " MEDIAL ZIPPER"
        Else
            sZipperText = " LATERAL ZIPPER"
        End If
        Dim sText As String = Trim(LGLEGDIA1.fnInchesToText(nZipLength)) + Chr(34) + sZipperText
        If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
            sText = fnDisplayToCM(nZipLength) + sZipperText
        End If
        PR_DrawTextAsSymbol(xyBlock, sZipperID, sText, 0)
        SetDBData("ID", sZipperID)
        SetDBData("Zipper", "1")
        SetDBData("Data", sText)

        ''// Reset And exit
        ''Execute("menu", "SetLayer", Table("find", "layer", "1")) ;
        MsgBox("Zipper drawing Complete", 0, "Waist Chap Ankle Zipper")

    End Sub
End Class

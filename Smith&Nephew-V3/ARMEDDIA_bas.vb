Option Strict Off
Option Explicit On
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports VB = Microsoft.VisualBasic
Module ARMEDDIA1
	'* Windows API Functions Declarations
	'    Declare Function GetWindow Lib "User" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
	'    Declare Function GetWindowText Lib "User" (ByVal hwnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
	'    Declare Function GetWindowTextLength Lib "User" (ByVal hwnd As Integer) As Integer
	'    Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
	'    Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)
	'
	'   'Constants used by GetWindow
	'    Const GW_CHILD = 5
	'    Const GW_HWNDFIRST = 0
	'    Const GW_HWNDLAST = 1
	'    Const GW_HWNDNEXT = 2
	'    Const GW_HWNDPREV = 3
	'    Const GW_OWNER = 4
	
	Public MainForm As armeddia
	
	Structure XY
		Dim X As Double
		Dim y As Double
	End Structure
	
	Structure curve
		Dim n As Short
		<VBFixedArray(100)> Dim X() As Double
		<VBFixedArray(100)> Dim y() As Double
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			'UPGRADE_WARNING: Lower bound of array X was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim X(100)
			'UPGRADE_WARNING: Lower bound of array y was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim y(100)
		End Sub
	End Structure
	
	'Globals
	Public g_nUnitsFac As Double
	
	Public g_iGms(17) As Short
	Public g_iRed(17) As Short
	Public g_iMMs(17) As Short
	Public g_nLengths(17) As Double
	
	Public g_iGmsInit(17) As Short
	Public g_iRedInit(17) As Short
	Public g_iMMsInit(17) As Short
	Public g_nLengthsInit(17) As Double
	Public g_iChanged(17) As Short
	Public g_iVertexMap(17) As Short 'Maps length to profile vertex
	
	Public g_iFirstTape As Short
	Public g_iLastTape As Short

    Public g_ReDrawn As Boolean

    'Profile Globals
    Public xyProfile(200) As XY
	Public xyOriginal(200) As XY
	Public g_iProfile As Short
	Public xyOtemplate As XY
	Public g_iFirstEditableVertex As Short
	Public g_iFirstEditTape As Short
	Public g_iLastEditTape As Short
	Public g_ShortArm As Short
	Public g_iElbowTape As Short
	Public g_NoElbowTape As Short
	Public g_sOriginalContracture As String
	Public g_sOriginalLining As String
	
	'Globals set by FN_Open
	Public cc As Object 'Comma
	Public QQ As Object 'Quote
	Public NL As Object 'Newline
	Public fNum As Short 'Macro file number
	Public QCQ As Object 'Quote Comma Quote
	Public QC As Object 'Quote Comma
	Public CQ As Object 'Comma Quote
	
	
	Public g_sPathJOBST As String
	
	Public g_nCurrTextHt As Object
	Public g_nCurrTextAspect As Object
	Public g_nCurrTextHorizJust As Object
	Public g_nCurrTextVertJust As Object
	Public g_nCurrTextFont As Object
	
	Public g_sFileNo As String
	Public g_sSide As String
	Public g_sPatient As String
	Public g_sCurrentLayer As String
	
	Public g_iWristNo As Short
	Public g_iPalmNo As Short
	
	Public g_nPalmWristDist As Double
    Public g_iModulus As Short
    Public g_sTempArmHandle As String
    Public g_sOriginalCurveHandle As String

    'POWERNET and BOBINNET Modulus charts
    'This format was initially created to allow DRAFIX to index into a string,
    'to mimic an array
    'This format is retained here for compatibility.
    Structure fabric
		<VBFixedArray(18)> Dim Modulus() As String
		'UPGRADE_NOTE: Conversion was upgraded to Conversion_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		<VBFixedArray(18)> Dim Conversion_Renamed() As String
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim Modulus(18)
		End Sub
	End Structure
	
	'UPGRADE_WARNING: Arrays in structure g_MATERIAL may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public g_MATERIAL As fabric
	
	'Flags
	Public g_Gauntlet As Short
	Public g_sID As String
	Public g_sStyle As String
	
	
	
	
	Function fnDisplayToInches(ByVal nDisplay As Double) As Double
		'This function takes the value given and converts it
		'into a decimal version in inches, rounded to the nearest eighth
		'of an inch.
		'
		'Input:-
		'        nDisplay is the value as input by the operator in the
		'        dialog.
		'        The convention is that, Metric dimensions use the decimal
		'        point to indicate the division between CMs and mms
		'        ie 7.6 = 7 cm and 6 mm.
		'        Whereas the decimal point for imperial measurements indicates
		'        the division between inches and eighths
		'        ie 7.6 = 7 inches and 6 eighths
		'Globals:-
		'        g_nUnitsFac = 1       => nDisplay in Inches
		'        g_nUnitsFac = 10/25.5 => nDisplay in CMs
		'Returns:-
		'        Single, Inches rounded to the nearest eighth (0.125)
		'        -1,     on conversion error.
		'
		'Errors:-
		'        The returned value is usually +ve. Unless it can't
		'        be sucessfully converted to inches.
		'        Eg 7.8 is an invalid number if g_nUnitsFac = 1
		'
		'                            WARNING
		'                            ~~~~~~~
		'In most cases the input is a +ve number.  This function will handle a
		'-ve number but in this case the error checking is invalid.  This
		'is done to provide a general conversion tool.  Where the input is
		'likley to be -ve then the calling subroutine or function should check
		'the sensibility of the returned value for that specific case.
		'
		
		Dim iInt, iSign As Short
		Dim nDec As Double
		'retain sign
		iSign = System.Math.Sign(nDisplay)
		nDisplay = System.Math.Abs(nDisplay)
		
		'Simple case where Units are CM
		If g_nUnitsFac <> 1 Then
			fnDisplayToInches = fnRoundInches(nDisplay * g_nUnitsFac) * iSign
			Exit Function
		End If
		
		'Imperial units
		iInt = Int(nDisplay)
		nDec = nDisplay - iInt
		'Check that conversion is possible (return -1 if not)
		If nDec > 0.8 Then
			fnDisplayToInches = -1
		Else
			fnDisplayToInches = fnRoundInches(iInt + (nDec * 0.125 * 10)) * iSign
		End If
		
	End Function
	
	
	
	Function fnGetNumber(ByVal sString As String, ByRef iIndex As Short) As Double
		'Function to return as a numerical value the iIndexth item in a string
		'that uses blanks (spaces) as delimiters.
		'EG
		'    sString = "12.3 65.1 45"
		'    fnGetNumber( sString, 2) = 65.1
		'
		'If the iIndexth item is not found then return -1 to indicate an error.
		'This assumes that the string will not be used to store -ve numbers.
		'Indexing starts from 1
		
		Dim ii, iPos As Short
		Dim sItem As String
		
		'Initial error checking
		sString = Trim(sString) 'Remove leading and trailing blanks
		
		If Len(sString) = 0 Then
			fnGetNumber = -1
			Exit Function
		End If
		
		'Prepare string
		sString = sString & " " 'Trailing blank as stopper for last item
		
		'Get iIndexth item
		For ii = 1 To iIndex
			iPos = InStr(sString, " ")
			If ii = iIndex Then
				sString = Left(sString, iPos - 1)
				fnGetNumber = Val(sString)
				Exit Function
			Else
				sString = LTrim(Mid(sString, iPos))
				If Len(sString) = 0 Then
					fnGetNumber = -1
					Exit Function
				End If
			End If
		Next ii
		
		'The function should have exited befor this, however just in case
		'(iIndex = 0) we indicate an error,
		fnGetNumber = -1
		
	End Function
	
	Function fnInchesToText(ByRef nInches As Double) As String
		'Function returns a decimal value in inches as a string
		'
		Dim nPrecision, nDec As Double
		Dim iInt, iEighths As Short
		Dim sString As String
		nPrecision = 0.125
		
		'Split into decimal parts
		iInt = Int(nInches)
		nDec = nInches - iInt
		If nDec <> 0 Then 'Avoid overflow
			iEighths = Int(nDec / nPrecision)
		Else
			iEighths = 0
		End If
		
		'Format string
		If iInt <> 0 Then
			sString = LTrim(Str(iInt))
		Else
			sString = "  "
		End If
		If iEighths <> 0 Then
			Select Case iEighths
				Case 2, 6
					sString = sString & "-" & LTrim(Str(iEighths / 2)) & "/4"
				Case 4
					sString = sString & "-" & "1/2"
				Case Else
					sString = sString & "-" & LTrim(Str(iEighths)) & "/8"
			End Select
		Else
			sString = sString & "   "
		End If
		
		'Return formatted string
		fnInchesToText = sString
		
	End Function
	
	Function fnRoundInches(ByVal nNumber As Double) As Double
		'Function to return the rounded value in decimal inches
		'returns to the nearest eighth (0.125)
		'E.G.
		'    5.67         = 5 inches and 0.67 inches
		'                   0.67 / 0.125 = 5.36 eighths
		'                   5.36 eighths = 5 eighths (rounded to nearest eighth)
		'    5.67         = 5 inches and 5 eighths
		'    5.67         = 5 + ( 5 * 0.125)
		'    5.67         = 6.625 inches
		'
		
		Dim iInt, iSign As Short
		Dim nPrecision, nDec As Double
		
		'Return 0 if input is Zero
		If nNumber = 0 Then
			fnRoundInches = 0
			Exit Function
		End If
		
		'Set precision
		nPrecision = 0.125
		
		'Break input into components
		iSign = System.Math.Sign(nNumber)
		nNumber = System.Math.Abs(nNumber)
		iInt = Int(nNumber)
		nDec = nNumber - iInt
		
		'Get decimal part in precision units
		If nDec <> 0 Then
			nDec = nDec / nPrecision 'Avoid overflow
		End If
		nDec = round(nDec)
		
		'Return value
		fnRoundInches = (iInt + (nDec * nPrecision)) * iSign
		
	End Function
	
	Sub PR_LoadFabricFromFile(ByRef MATERIAL As fabric, ByRef sFileName As String)
		'Procedure to load the MATERIAL conversion chart from file
		'N.B. File opening Errors etc. are not handled (so tough titty!)
		
		Dim fNum, ii As Short
		fNum = FreeFile()
        FileOpen(fNum, sFileName, VB.OpenMode.Input)
        ii = 0
		Do Until EOF(fNum)
			Input(fNum, MATERIAL.Modulus(ii))
			Input(fNum, MATERIAL.Conversion_Renamed(ii))
			ii = ii + 1
		Loop 
		FileClose(fNum)
		
	End Sub
	
	Function round(ByVal nNumber As Double) As Short
		'Fuction to return the rounded value of a decimal number
		'E.G.
		'    round(1.35)  = 1
		'    round(1.55)  = 2
		'    round(2.50)  = 3
		'    round(-2.50) = -3
		'    round(0)     = 0
		'
		
		Dim iInt, iSign As Short
		
		'Avoid extra work. Return 0 if input is 0
		If nNumber = 0 Then
			round = 0
			Exit Function
		End If
		
		'Split input
		iSign = System.Math.Sign(nNumber)
		nNumber = System.Math.Abs(nNumber)
		iInt = Int(nNumber)
		
		'Effect rounding
		If (nNumber - iInt) >= 0.5 Then
			round = (iInt + 1) * iSign
		Else
			round = iInt * iSign
		End If
		
	End Function

    Sub Select_Text_In_Box(ByRef Text_Box_Name As System.Windows.Forms.Control)
        Text_Box_Name.Focus()
        'UPGRADE_WARNING: Couldn't resolve default property of object Text_Box_Name.SelStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ''Text_Box_Name.SelStart = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object Text_Box_Name.SelLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ''Text_Box_Name.SelLength = Len(Text_Box_Name.Text)
    End Sub
    Sub PR_DrawPoly(ByRef Profile As curve)
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
                '    acPoly.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyOtemplate.X, xyOtemplate.y, 0)))
                'End If
                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acPoly)
                acTrans.AddNewlyCreatedDBObject(acPoly, True)
                Dim acRegAppTbl As RegAppTable
                acRegAppTbl = acTrans.GetObject(acCurDb.RegAppTableId, OpenMode.ForRead)
                Dim acRegAppTblRec As RegAppTableRecord
                Dim appName As String = g_sID & "Contracture"
                Dim xdataStr As String = g_sID & "Contracture"
                If acRegAppTbl.Has(appName) = False Then
                    acRegAppTblRec = New RegAppTableRecord
                    acRegAppTblRec.Name = appName
                    acRegAppTbl.UpgradeOpen()
                    acRegAppTbl.Add(acRegAppTblRec)
                    acTrans.AddNewlyCreatedDBObject(acRegAppTblRec, True)
                End If
                Using rb As New ResultBuffer
                    rb.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, appName))
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, xdataStr))
                    acPoly.XData = rb
                End Using
            End Using

            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_DrawText(ByRef sText As Object, ByRef xyInsert As XY, ByRef nHeight As Object, ByRef nAngle As Double)
        'To the DRAFIX macro file (given by the global fNum).
        'Write the syntax to draw TEXT at the given height.
        '
        'For this to work it assumes that the following DRAFIX variables
        'are defined
        '    XY      xyStart
        '
        'Note:-
        '    fNum, CC, QQ, NL, g_nCurrTextAspect are globals initialised by FN_Open
        '
        '
        Dim nWidth As Object
        nWidth = nHeight * g_nCurrTextAspect
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
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)

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
    Sub PR_MakeXY(ByRef xyReturn As XY, ByRef X As Double, ByRef y As Double)
        'Utility to return a point based on the X and Y values
        'given
        xyReturn.X = X
        xyReturn.y = y
    End Sub
    'To Draw Spline
    Sub PR_DrawSpline(ByRef PointCollection As Point3dCollection, Optional ByVal bIsSetXDATA As Boolean = False)
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
                If bIsSetXDATA = True Then
                    Dim acRegAppTbl As RegAppTable
                    acRegAppTbl = acTrans.GetObject(acCurDb.RegAppTableId, OpenMode.ForRead)
                    Dim acRegAppTblRec As RegAppTableRecord
                    Dim appName As String = "ProfileID"
                    Dim xdataStr As String = "ARM" & ARMDIA1.g_sID & "Profile"
                    If acRegAppTbl.Has(appName) = False Then
                        acRegAppTblRec = New RegAppTableRecord
                        acRegAppTblRec.Name = appName
                        acRegAppTbl.UpgradeOpen()
                        acRegAppTbl.Add(acRegAppTblRec)
                        acTrans.AddNewlyCreatedDBObject(acRegAppTblRec, True)
                    End If
                    Using rb As New ResultBuffer
                        rb.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, appName))
                        rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, xdataStr))
                        acSpline.XData = rb
                    End Using
                End If
                g_sTempArmHandle = acSpline.Handle.ToString()
            End Using
            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Sub PR_DrawTextCenter(ByRef sText As Object, ByRef xyInsert As XY, ByRef nHeight As Object, ByRef nAngle As Double)
        Dim nWidth As Object
        nWidth = nHeight * g_nCurrTextAspect

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
                acText.Position = New Point3d(xyInsert.X + xyInsertion.X, xyInsert.y + xyInsertion.y, 0)
                acText.Height = nHeight
                acText.TextString = sText
                acText.Rotation = nAngle
                acText.HorizontalMode = TextHorizontalMode.TextMid
                acText.AlignmentPoint = New Point3d(xyInsert.X, xyInsert.y, 0)
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
    Sub PR_DrawCircle(ByRef xyCen As XY, ByRef nRadius As Object)
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

            '' Create a circle that is at 2,3 with a radius of 4.25
            Using acCirc As Circle = New Circle()
                acCirc.Center = New Point3d(xyCen.X, xyCen.y, 0)
                acCirc.Radius = nRadius
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acCirc.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyCen.X, xyCen.y, 0)))
                End If

                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acCirc)
                acTrans.AddNewlyCreatedDBObject(acCirc, True)
            End Using

            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
End Module
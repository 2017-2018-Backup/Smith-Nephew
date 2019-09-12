Option Strict Off
Option Explicit On
Module MESHCALC
	'Module:    MESHCALC.MAK
	'Purpose:   Mesh axilla calculation
	'
	'Projects:  1. BODYDRAW.MAK
	'           2. MESHDRAW.MAK
	'
	'Version:   1.00
	'Date:      5.Nov.1997
	'Author:    Gary George
	'           © C-Gem Ltd.
	'
	'-------------------------------------------------------
	'REVISIONS:
	'Date       By      Action
	'-------------------------------------------------------
	'
	'Notes:-
	'
	'
	'
	
	
	Public Structure curve
		Dim n As Short
		<VBFixedArray(100)> Dim x() As Double
		<VBFixedArray(100)> Dim y() As Double
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			'UPGRADE_WARNING: Lower bound of array x was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim x(100)
			'UPGRADE_WARNING: Lower bound of array y was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim y(100)
		End Sub
	End Structure
	
	Public Structure XY
		Dim x As Double
		Dim y As Double
	End Structure
	
	Public Structure BiArc
		Dim xyStart As XY
		Dim xyTangent As XY
		Dim xyEnd As XY
		Dim xyR1 As XY
		Dim xyR2 As XY
		Dim nR1 As Double
		Dim nR2 As Double
	End Structure
	
	Public xyAxilla As XY
	
	Public Const ONE_and_THREE_QUARTER_GUSSET As Double = 2.58
	Public Const BOYS_GUSSET As Double = 2.8
	
	Function FN_CalcAxillaMesh(ByRef xyRaglanStart As XY, ByRef xyRaglanEnd As XY, ByRef xyMeshStart As XY, ByRef nFlip As Short, ByRef nMeshLength As Double, ByRef nAlongRaglan As Double, ByRef MeshProfile As BiArc, ByRef MeshSeam As BiArc) As Short
		'Subroutine to calculate a mesh w.r.t Mesh axilla in the
		'Vest, Body-suit and their respective sleeves
		
		'Due to problems with the CALC BiARC function we calculate all for the lower
		'half and then flip the result
		'A cheat but what the hell
		
		Dim aDeltaEnd, aEnd, aInitialStart, aStart, aAngle, aInitialEnd, aMeshEnd, aStartDelta As Double
		Dim nEndLength, nStartLength, nStep As Double
		Dim nRadiusTol, nLength, nTol, nSeam As Double
		Dim bMeshFound, ii As Short
		Dim aAxis, iSign, aEndFactor As Double
		Dim xyMeshEnd, xyTmp As XY
		'UPGRADE_WARNING: Arrays in structure cDummyCurve may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim cDummyCurve As curve
		
		'intially start as False
		FN_CalcAxillaMesh = False
		
		'Check for silly data
		If xyRaglanStart.x = xyRaglanEnd.x And xyRaglanStart.y = xyRaglanEnd.y Then Exit Function
		If xyMeshStart.x = xyRaglanEnd.x And xyMeshStart.y = xyRaglanEnd.y Then Exit Function
		If xyRaglanStart.x = xyMeshStart.x And xyRaglanStart.y = xyMeshStart.y Then Exit Function
		
		If ARMDIA1.FN_CalcLength(xyMeshStart, xyRaglanStart) >= nMeshLength Then Exit Function
		
		
		'initialize variables
		'    nTol = .03125
		nTol = 0.04
		bMeshFound = False
		nStep = -INCH1_16
		
		If nAlongRaglan > 0 Then
			nStartLength = nAlongRaglan
			nEndLength = nAlongRaglan 'do at least Once
			aDeltaEnd = 160 'Must find
		Else
			nLength = ARMDIA1.FN_CalcLength(xyMeshStart, xyRaglanStart)
			nEndLength = nLength + INCH3_16
			nStartLength = nLength * 2.5 '
			
			aDeltaEnd = 90
			nStep = INCH1_16
			If nMeshLength = ONE_and_THREE_QUARTER_GUSSET Then
				nEndLength = nLength * 2.5 '
				nStartLength = 1.25 '
			Else
				nEndLength = nLength * 2.5 '
				nStartLength = 1.5 '
			End If
			
		End If
		
		If My.Application.Info.Title <> "meshdraw" Then
			'Body suit Mesh
			aAxis = ARMDIA1.FN_CalCAngle(xyMeshStart, xyRaglanStart)
			aStartDelta = 40
			'        nRadiusTol = .4
			nRadiusTol = 0
			aInitialStart = 265
		Else
			'Sleeve Mesh
			aAxis = ARMDIA1.FN_CalCAngle(xyMeshStart, xyAxilla)
			aStartDelta = 60
			aInitialStart = 280
			nRadiusTol = 0
		End If
		
		For aStart = aInitialStart To (aInitialStart + aStartDelta) Step 1
			'PR_CalcPolar xyMeshStart, aStart, 1, xyTmp
			'PR_DrawMarker xyTmp
			
			For nLength = nStartLength To nEndLength Step nStep
				'Get point on raglan
				PR_CalcRaglan(xyRaglanStart, xyRaglanEnd, nFlip, cDummyCurve, xyMeshEnd, aInitialEnd, nLength)
				' PR_CalcPolar xyMeshEnd, aInitialEnd, 3, xyTmp
				' PR_DrawLine xyTmp, xyMeshEnd
				
				'Flip the point and angle
				If nFlip Then
					'mirror in the line (xyMeshStart, xyRaglanStart)
					ARMDIA1.PR_CalcPolar(xyMeshStart, (360 - ARMDIA1.FN_CalCAngle(xyMeshStart, xyMeshEnd) + aAxis), ARMDIA1.FN_CalcLength(xyMeshStart, xyMeshEnd), xyMeshEnd)
					aMeshEnd = ((360 - aInitialEnd) + aAxis) + 180
				Else
					aMeshEnd = (aInitialEnd + 180)
				End If
				
				'End angle iteration - Get biarc curve that fits the given length to +/- nTol
				'Hold start change end
				For aEnd = (aMeshEnd - (aDeltaEnd / 2)) To (aMeshEnd + (aDeltaEnd / 2)) Step 1
					' PR_CalcPolar xyMeshEnd, aEnd, 1, xyTmp
					' PR_DrawMarker xyTmp
					
					If FN_FitBiArcAtGivenLength(xyMeshStart, aStart, xyMeshEnd, aEnd, nMeshLength, nTol, nRadiusTol, MeshProfile) Then
						bMeshFound = True
						Exit For
					End If
				Next aEnd
				If bMeshFound Then
					FN_CalcAxillaMesh = True
					nAlongRaglan = nLength
					Exit For
				End If
				
			Next nLength
			
			If bMeshFound Then Exit For
			
		Next aStart
		
		'Don't calculate the seam unless we have a mesh
		If Not bMeshFound Then Exit Function
		
		'Calculate seam line
		nSeam = 0.125
		
		'Calculate MeshSeam based on MeshProfile
		If nFlip Then
			'mirror in the line (xyMeshStart, xyRaglanStart)
			ARMDIA1.PR_CalcPolar(xyMeshStart, (360 - ARMDIA1.FN_CalCAngle(xyMeshStart, MeshProfile.xyEnd) + aAxis), ARMDIA1.FN_CalcLength(xyMeshStart, MeshProfile.xyEnd), MeshProfile.xyEnd)
			ARMDIA1.PR_CalcPolar(xyMeshStart, (360 - ARMDIA1.FN_CalCAngle(xyMeshStart, MeshProfile.xyTangent) + aAxis), ARMDIA1.FN_CalcLength(xyMeshStart, MeshProfile.xyTangent), MeshProfile.xyTangent)
			ARMDIA1.PR_CalcPolar(xyMeshStart, (360 - ARMDIA1.FN_CalCAngle(xyMeshStart, MeshProfile.xyR1) + aAxis), ARMDIA1.FN_CalcLength(xyMeshStart, MeshProfile.xyR1), MeshProfile.xyR1)
			ARMDIA1.PR_CalcPolar(xyMeshStart, (360 - ARMDIA1.FN_CalCAngle(xyMeshStart, MeshProfile.xyR2) + aAxis), ARMDIA1.FN_CalcLength(xyMeshStart, MeshProfile.xyR2), MeshProfile.xyR2)
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object MeshSeam. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MeshSeam = MeshProfile
		
		'Get seam intersection on raglan
		nLength = nLength - nSeam
		PR_CalcRaglan(xyRaglanStart, xyRaglanEnd, nFlip, cDummyCurve, MeshSeam.xyEnd, aInitialEnd, nLength)
		
		'Calculate Start, End and Tangent
		aAngle = ARMDIA1.FN_CalCAngle(MeshProfile.xyR1, MeshProfile.xyStart)
		MeshSeam.nR1 = MeshProfile.nR1 - nSeam
		ARMDIA1.PR_CalcPolar(MeshProfile.xyR1, aAngle, MeshSeam.nR1, MeshSeam.xyStart)
		
		aAngle = ARMDIA1.FN_CalCAngle(MeshProfile.xyR1, MeshProfile.xyTangent)
		ARMDIA1.PR_CalcPolar(MeshProfile.xyR1, aAngle, MeshSeam.nR1, MeshSeam.xyTangent)
		MeshSeam.nR2 = MeshProfile.nR2 - nSeam
		
		'PR_DrawMarker MeshSeam.xyR2
		'PR_DrawText "MeshSeam.xyStart", MeshSeam.xyStart, .1
		
	End Function
	
	
	Function FN_FitBiArcAtGivenLength(ByRef xyStart As XY, ByRef aStart As Double, ByRef xyEnd As XY, ByRef aEnd As Double, ByRef nLength As Double, ByRef nLengthTol As Double, ByRef nRadiusTol As Double, ByRef Mesh As BiArc) As Short
		
		Dim nArc, aA1, aA2, aArc As Double
		
		FN_FitBiArcAtGivenLength = False
		If BDYUTILS.FN_BiArcCurve(xyStart, aStart, xyEnd, aEnd, Mesh) Then
			'Calculate Length
			'First arc
			aA1 = ARMDIA1.FN_CalCAngle(Mesh.xyR1, Mesh.xyStart)
			aA2 = ARMDIA1.FN_CalCAngle(Mesh.xyR1, Mesh.xyTangent)
			aArc = aA2 - aA1
			nArc = (aArc * (PI / 180)) * Mesh.nR1
			'Second arc
			aA1 = ARMDIA1.FN_CalCAngle(Mesh.xyR2, Mesh.xyTangent)
			aA2 = ARMDIA1.FN_CalCAngle(Mesh.xyR2, Mesh.xyEnd)
			aArc = aA2 - aA1
			nArc = nArc + ((aArc * (PI / 180)) * Mesh.nR2)
			
			'Check that the length is within the tolerence
			If nArc > (nLength - nLengthTol) And nArc < (nLength + nLengthTol) Then
				FN_FitBiArcAtGivenLength = True
				If nRadiusTol > 0 And System.Math.Abs(Mesh.nR1 - Mesh.nR2) > nRadiusTol Then FN_FitBiArcAtGivenLength = False
			End If
			
		End If
		
	End Function
	
	Sub PR_CalcRaglan(ByRef xyStart As XY, ByRef xyEnd As XY, ByRef Flip As Short, ByRef Raglan As curve, ByRef xyOpen As XY, ByRef aOpen As Double, ByRef nRadiusForRequiredPoint As Double)
		'Subroutine to draw a vest raglan curve between the two points
		'given
		'NB
		' nRadiusForRequiredPoint is used to get the point xyOpen
		' if nRadiusForRequiredPoint = -1 then xyOpen is at 1/3 the distance along the curve
		' if nRadiusForRequiredPoint > 0  then xyOpen is at the radius nRadiusForRequiredPoint
		'
		Static nSegLen(100) As Double
		Static nSegAngle(100) As Double
		Static iSegments As Short
		Static sFileName As String
		Static fFileNum As Short
		Static xy2, xy0, xy1, xy4 As XY
		Static nLength, aPrevAngle, aAngle, nTraversedLen, nOpenLength As Double
		Static bOpenFound As Short
		Static aAxis, aRaglan, nRadius As Double
		Static ii As Short
		
		'On an error the Raglan Curve is returned as a minimum of 2 points
		'these being the given start and end points.
		Raglan.n = 2
		Raglan.x(1) = xyStart.x
		Raglan.y(1) = xyStart.y
		Raglan.x(2) = xyEnd.x
		Raglan.y(2) = xyEnd.y
		
		'Check for silly data
		If xyStart.y = xyEnd.y And xyStart.x = xyEnd.x Then Exit Sub
		
		If g_sSleeveType = "Sleeveless" Then
			sFileName = g_sPathJOBST & "\TEMPLTS\BODYCURV.DAT"
		Else
			sFileName = g_sPathJOBST & "\TEMPLTS\VESTCURV.DAT"
		End If
		'Check to see if the raglan curve has been loaded from file,
		'so we don't load it twice
		If iSegments = 0 Then
			fFileNum = FreeFile
			If FileLen(sFileName) = 0 Then
				MsgBox("Can't open the file " & sFileName & ".  Unable to draw the raglan curve.", 48)
				Exit Sub
			End If
			
			FileOpen(fFileNum, sFileName, OpenMode.Input)
			Do While Not EOF(fFileNum)
				iSegments = iSegments + 1
				Input(fFileNum, nSegLen(iSegments))
				Input(fFileNum, nSegAngle(iSegments))
			Loop 
			FileClose(fFileNum)
		End If
		
		'Using the given points loop through the raglan curve to establish the
		'raglan angle.
		'Note:
		'This code is a little bit cumbersome but it is based on DRAFIX macro
		'code that is known to work and is production accepted.
		'
		nRadius = ARMDIA1.FN_CalcLength(xyStart, xyEnd)
		
		aRaglan = -1
		xy1.x = 0 : xy1.y = 0
		xy0.x = 0 : xy0.y = 0
		aPrevAngle = 0
		aAngle = nSegAngle(1)
		nLength = nSegLen(1)
		nTraversedLen = 0 'For use later
		
		For ii = 2 To iSegments
			aAngle = aAngle + aPrevAngle
			ARMDIA1.PR_CalcPolar(xy1, aAngle, nLength, xy2)
			If BDYUTILS.FN_CirLinInt(xy1, xy2, xy0, nRadius, xy4) Then
				aRaglan = ARMDIA1.FN_CalCAngle(xy0, xy4)
				nTraversedLen = nTraversedLen + ARMDIA1.FN_CalcLength(xy1, xy4)
				Exit For
			End If
			nTraversedLen = nTraversedLen + ARMDIA1.FN_CalcLength(xy1, xy2)
			'UPGRADE_WARNING: Couldn't resolve default property of object xy1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xy1 = xy2
			aPrevAngle = aAngle
			aAngle = nSegAngle(ii)
			nLength = nSegLen(ii)
		Next ii
		
		'Check that the intesection has occured
		If aRaglan < 0 Then
			MsgBox("Can't draw the draw the raglan curve with this data.", 48)
			Exit Sub
		End If
		
		'Create the curve given the angle above and the two point arguments
		'
		nOpenLength = nTraversedLen / 3
		nTraversedLen = 0
		bOpenFound = False
		
		aAngle = nSegAngle(1)
		nLength = nSegLen(1)
		aAxis = ARMDIA1.FN_CalCAngle(xyStart, xyEnd)
		aPrevAngle = aAxis - aRaglan
		'UPGRADE_WARNING: Couldn't resolve default property of object xy1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xy1 = xyStart 'NB Raglan.X(1) and Raglan.Y(1), set at begining to xyStart
		Raglan.x(1) = xyStart.x
		Raglan.y(1) = xyStart.y
		For ii = 2 To iSegments
			aAngle = aAngle + aPrevAngle
			ARMDIA1.PR_CalcPolar(xy1, aAngle, nLength, xy2)
			If BDYUTILS.FN_CirLinInt(xy1, xy2, xyStart, nRadius, xy4) Then
				Raglan.x(ii) = xyEnd.x
				Raglan.y(ii) = xyEnd.y
				Exit For
			End If
			If nRadiusForRequiredPoint > 0 Then
				If BDYUTILS.FN_CirLinInt(xy1, xy2, xyStart, nRadiusForRequiredPoint, xy4) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object xyOpen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					xyOpen = xy4
					aOpen = ARMDIA1.FN_CalCAngle(xy1, xy2)
					If Flip Then
						aOpen = (360 - (aOpen - aAxis) + aAxis) + 90
					Else
						aOpen = aOpen - 90
					End If
				End If
			End If
			Raglan.x(ii) = xy2.x
			Raglan.y(ii) = xy2.y
			If nTraversedLen + nLength >= nOpenLength And Not bOpenFound And nRadiusForRequiredPoint <= 0 Then
				bOpenFound = True
				aOpen = ARMDIA1.FN_CalCAngle(xy1, xy2)
				ARMDIA1.PR_CalcPolar(xy1, aOpen, (nTraversedLen + nLength) - nOpenLength, xyOpen)
				If Flip Then
					aOpen = (360 - (aOpen - aAxis) + aAxis) + 90
				Else
					aOpen = aOpen - 90
				End If
			End If
			nTraversedLen = nTraversedLen + nLength
			'UPGRADE_WARNING: Couldn't resolve default property of object xy1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xy1 = xy2
			aPrevAngle = aAngle
			aAngle = nSegAngle(ii)
			nLength = nSegLen(ii)
			
		Next ii
		Raglan.n = ii
		
		'If flip is true then mirror Raglan curve in the line (xyStart, xyEnd)
		'The angle of which is given by aAxis
		If Flip Then
			For ii = 1 To Raglan.n
				ARMDIA1.PR_MakeXY(xy1, Raglan.x(ii), Raglan.y(ii))
				nLength = ARMDIA1.FN_CalcLength(xyStart, xy1)
				aAngle = ARMDIA1.FN_CalCAngle(xyStart, xy1)
				aAngle = (360 - (aAngle - aAxis) + aAxis)
				ARMDIA1.PR_CalcPolar(xyStart, aAngle, nLength, xy2)
				Raglan.x(ii) = xy2.x
				Raglan.y(ii) = xy2.y
			Next ii
			nLength = ARMDIA1.FN_CalcLength(xyStart, xyOpen)
			aAngle = ARMDIA1.FN_CalCAngle(xyStart, xyOpen)
			aAngle = (360 - (aAngle - aAxis) + aAxis)
			ARMDIA1.PR_CalcPolar(xyStart, aAngle, nLength, xyOpen)
			
		End If
		
		
	End Sub
End Module
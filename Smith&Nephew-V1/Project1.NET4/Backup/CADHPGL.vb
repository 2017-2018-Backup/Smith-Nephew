Option Strict Off
Option Explicit On
Module CADHPGL
	'Project:   CADGLOVE.VBP
	'File:      CADHPGL.BAS
	'Purpose:   HPGL Interpreter and DRAFIX translator
	'           module.
	'
	'Version:   3.01
	'Date:      17.Jan.96
	'Author:    Gary George
	'
	'-------------------------------------------------------
	'REVISIONS:
	'Date       By      Action
	'-------------------------------------------------------
	'07 Dec 98  GG      Move the interaction with the CAD
	'                   Glove external programme to the module
	'                   Kevin
	'Notes:-
	'
	
	Private miFileHPGL As Short
	'UPGRADE_WARNING: Lower bound of array Value was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Private Value(10) As Object
	
	'The TextLabels are use to identify points
	'
	Structure TextLabel
		Dim Found As Short
		Dim sString As String 'If any
		Dim X As Double
		Dim y As Double
	End Structure
	
	'UPGRADE_WARNING: Lower bound of array g_Labels was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim g_Labels(30) As TextLabel
	
	
	
	
	'
	'Notes:-
	'
	'   This implementation of a stack as an array is lifted directly
	'   from
	'       Data Structures and Program Design
	'       Robert L. Kruse, 1984
	'       2nd Edition, Prentice Hall International Editions
	'
	'
	
	Const MAXSTACK As Short = 100
	
	Structure stack
		Dim Top As Short
		<VBFixedArray(MAXSTACK)> Dim X() As Double
		<VBFixedArray(MAXSTACK)> Dim y() As Double
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			'UPGRADE_WARNING: Lower bound of array X was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim X(MAXSTACK)
			'UPGRADE_WARNING: Lower bound of array y was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim y(MAXSTACK)
		End Sub
	End Structure
	
	
	'UPGRADE_WARNING: Arrays in structure S may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Public S As stack
	
	
	Sub ClearStack()
		S.Top = 0
	End Sub
	
	Sub DrawMarker()
		Dim xyStart As XY
		
		'Draw a Marker at the top of the current stack point
		
		Dim Xlast, Ylast As Double
		
		'Return if empty
		If EmptyStack() Then Exit Sub
		
		'Get last point
		StackCurrent(Xlast, Ylast)
		ARMDIA1.PR_MakeXY(xyStart, Xlast / ISCALE, Ylast / ISCALE)
		PR_RevisePoint(xyStart)
		
		BDYUTILS.PR_DrawMarker(xyStart)
		
	End Sub
	
	Sub DrawText(ByRef Text As Object)
		Dim xyStart As XY
		Dim Xstart, Ystart As Double
		
		StackCurrent(Xstart, Ystart)
		ARMDIA1.PR_MakeXY(xyStart, Xstart / ISCALE, Ystart / ISCALE)
		PR_RevisePoint(xyStart)
		
		ARMDIA1.PR_DrawText(Text, xyStart, 0.1)
		
	End Sub
	
	Function EmptyStack() As Short
		If S.Top = 0 Then EmptyStack = True Else EmptyStack = False
	End Function
	
	
	
	Function FullStack() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object FullStack. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If S.Top = MAXSTACK Then
			FullStack = True
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object FullStack. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FullStack = False
		End If
	End Function
	
	Sub PopStack(ByRef X As Double, ByRef y As Double)
		If S.Top = 0 Then
			MsgBox("Stack is Empty", 16, "Stack Operations")
			End
		Else
			X = S.X(S.Top)
			y = S.y(S.Top)
			S.Top = S.Top - 1
		End If
		
	End Sub
	
	Sub PR_AddNotesEtc()
		'This is a catchall procedure to draw everything that is not
		'part of the glove outline.
		'It uses the "hooks" created in the Labels "LB" statements in the
		'HPGL code.  It's ugly but it works!
		'
		Static ii As Short
		Static sText, sSymbol As String
		Static TmpX, TmpY As Double
		'UPGRADE_WARNING: Arrays in structure Reinforced may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Static xyStart, xyInsert, xyEnd As XY
		Static Reinforced As Curve
		Static xyPt As XY
		Static nYStep, nInsert, nMaxY, nMaxX, nMinX, nMinY, nXStep, nSpacing As Double
		Static sWorkOrder As String
		
		'Patient Details
		If g_Labels(6).Found = True Then
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextAngle = 0
			
			ARMDIA1.PR_SetTextData(1, 32, -1, -1, -1) 'Horiz-Left, Vertical-Bottom
			
			'Main details in Green on layer notes
			ARMDIA1.PR_SetLayer("Notes")
			
			ARMDIA1.PR_MakeXY(xyInsert, (g_Labels(6).X / ISCALE) - 1.25, (g_Labels(6).y / ISCALE) - 0.25)
			PR_RevisePoint(xyInsert)
			If CType(MainForm.Controls("txtWorkOrder"), Object).Text = "" Then
				sWorkOrder = "-"
			Else
				sWorkOrder = CType(MainForm.Controls("txtWorkOrder"), Object).Text
			End If
			sText = CType(MainForm.Controls("txtSide"), Object).Text & "\n" & CType(MainForm.Controls("txtPatientName"), Object).Text & "\n" & sWorkOrder & "\n" & Trim(Mid(CType(MainForm.Controls("cboFabric"), Object).Text, 4))
			ARMDIA1.PR_DrawText(sText, xyInsert, 0.125)
			
			
			'Other patient details in black on layer construct
			ARMDIA1.PR_SetLayer("Construct")
			
			ARMDIA1.PR_MakeXY(xyInsert, xyInsert.X + 0.8, xyInsert.y)
			sText = CType(MainForm.Controls("txtFileNo"), Object).Text & "\n" & CType(MainForm.Controls("txtDiagnosis"), Object).Text & "\n" & CType(MainForm.Controls("txtAge"), Object).Text & "\n" & CType(MainForm.Controls("txtSex"), Object).Text
			ARMDIA1.PR_DrawText(sText, xyInsert, 0.125)
			
		End If
		
		ARMDIA1.PR_SetLayer("Notes")
		ARMDIA1.PR_SetTextData(2, 16, -1, -1, -1)
		
		'Thumb and palm pieces
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextAngle = 90
		For ii = 7 To 8
			If g_Labels(ii).Found = True Then
				sText = CType(MainForm.Controls("txtSide"), Object).Text & "\n" & CType(MainForm.Controls("txtPatientName"), Object).Text & "\n" & CType(MainForm.Controls("txtWorkOrder"), Object).Text
				ARMDIA1.PR_MakeXY(xyInsert, g_Labels(ii).X / ISCALE, g_Labels(ii).y / ISCALE)
				PR_RevisePoint(xyInsert)
				ARMDIA1.PR_DrawText(sText, xyInsert, 0.125)
			End If
		Next ii
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextAngle = 0
		
		'Position Insert text Size w.r.t lowest web position
		nMaxX = 0
		nMinX = 1.23456789E+18
		nMaxY = 0
		nMinY = 1.23456789E+18
		For ii = 1 To 5
			If g_Labels(ii).Found = True Then
				'UPGRADE_WARNING: Couldn't resolve default property of object max(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				nMaxX = ARMDIA1.max(nMaxX, g_Labels(ii).X)
				'UPGRADE_WARNING: Couldn't resolve default property of object min(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				nMinX = BDYUTILS.min(nMinX, g_Labels(ii).X)
				'UPGRADE_WARNING: Couldn't resolve default property of object max(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				nMaxY = ARMDIA1.max(nMaxY, g_Labels(ii).y)
				'UPGRADE_WARNING: Couldn't resolve default property of object min(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				nMinY = BDYUTILS.min(nMinY, g_Labels(ii).y)
			End If
		Next ii
		
		'Get insert size
		sText = "Unknown "
		If g_Labels(9).Found = True Then
			'Use the given insert value
			'else use caculated value returned from the CAD-Glove DOS Programme
			'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!cboInsert.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If CType(MainForm.Controls("cboInsert"), Object).ListIndex > 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!cboInsert.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				g_iInsertSize = CType(MainForm.Controls("cboInsert"), Object).ListIndex
				nInsert = g_iInsertSize * 0.125
			Else
				nInsert = Val(Mid(g_Labels(9).sString, 3))
				If nInsert <> 0 Then g_iInsertSize = ARMDIA1.round(nInsert / 0.125)
			End If
			sText = Trim(ARMEDDIA1.fnInchestoText(nInsert))
			If Mid(sText, 1, 1) = "-" Then sText = Mid(sText, 2) 'Strip leading "-" sign
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sText = sText & "\" & QQ
		End If
		sText = "INSERT " & sText
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextAngle = 0
		ARMDIA1.PR_MakeXY(xyInsert, (nMaxX / ISCALE) + nInsert, (nMinY + ((nMaxY - nMinY) / 2)) / ISCALE)
		PR_RevisePoint(xyInsert)
		ARMDIA1.PR_DrawText(sText, xyInsert, 0.125)
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextAngle = 0
		
		'Elastic
		If Val(CType(MainForm.Controls("txtAge"), Object).Text) <= 10 And g_ExtendTo = GLOVE_NORMAL Then
			ARMDIA1.PR_MakeXY(xyInsert, g_Labels(6).X / ISCALE, (nMinY + ((nMaxY - nMinY) / 2)) / ISCALE)
			PR_RevisePoint(xyInsert)
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sText = "1/2\" & QQ & " Elastic"
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextAngle = 0
			ARMDIA1.PR_DrawText(sText, xyInsert, 0.125)
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextAngle = 0
		End If
		
		'Slanted inserts
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!chkSlantedInserts.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CType(MainForm.Controls("chkSlantedInserts"), Object).Value = 1 Then
			nInsert = nInsert - 0.125
			For ii = 2 To 4
				If g_Labels(ii).Found = True Then
					TmpX = g_Labels(ii).X / ISCALE
					TmpY = g_Labels(ii).y / ISCALE
					
					'Horizontal line
					ARMDIA1.PR_MakeXY(xyStart, TmpX, TmpY)
					PR_RevisePoint(xyStart)
					ARMDIA1.PR_MakeXY(xyEnd, TmpX + nInsert, TmpY)
					PR_RevisePoint(xyEnd)
					ARMDIA1.PR_DrawLine(xyStart, xyEnd)
					BDYUTILS.PR_AddDBValueToLast("ID", g_sID)
					
					'Vertical line
					ARMDIA1.PR_MakeXY(xyStart, TmpX + nInsert, TmpY + 0.125)
					PR_RevisePoint(xyStart)
					ARMDIA1.PR_MakeXY(xyEnd, TmpX + nInsert, TmpY - 0.125)
					PR_RevisePoint(xyEnd)
					ARMDIA1.PR_DrawLine(xyStart, xyEnd)
					BDYUTILS.PR_AddDBValueToLast("ID", g_sID)
					
				End If
			Next ii
		End If
		
		'Reinforced Palms
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!chkPalm.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CType(MainForm.Controls("chkPalm"), Object).Value = 1 Then
			For ii = 1 To 5
				If g_Labels(ii).Found = True Then
					Reinforced.n = Reinforced.n + 1
					ARMDIA1.PR_MakeXY(xyPt, (g_Labels(ii).X / ISCALE) + 0.25, g_Labels(ii).y / ISCALE)
					PR_RevisePoint(xyPt)
					Reinforced.X(Reinforced.n) = xyPt.X
					Reinforced.y(Reinforced.n) = xyPt.y
				End If
			Next ii
			
			'Draw reinforcing lines using Long Dash line type
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "SetStyle" & QC & "Table(" & QQ & "find" & QCQ & "style" & QCQ & "Long Dash" & QQ & "));")
			
			ARMDIA1.PR_DrawFitted(Reinforced)
			ARMDIA1.PR_MakeXY(xyStart, g_Labels(10).X / ISCALE, g_Labels(10).y / ISCALE)
			ARMDIA1.PR_MakeXY(xyEnd, g_Labels(11).X / ISCALE, g_Labels(11).y / ISCALE)
			PR_RevisePoint(xyStart)
			PR_RevisePoint(xyEnd)
			ARMDIA1.PR_DrawLine(xyStart, xyEnd)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "SetStyle" & QC & "Table(" & QQ & "find" & QCQ & "style" & QCQ & "bylayer" & QQ & "));")
			
			'Add REINFORCED text
			TmpX = (g_Labels(5).X + ((g_Labels(11).X - g_Labels(5).X) / 2)) / ISCALE
			TmpY = g_Labels(11).y / ISCALE
			
			ARMDIA1.PR_MakeXY(xyInsert, TmpX, TmpY)
			PR_RevisePoint(xyInsert)
			ARMDIA1.PR_DrawText("REINFORCED\nPALM", xyInsert, 0.125)
			
		End If
		
		'Reinforced Dorsal
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!chkDorsal.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Static nDorsalOffset As Double
		If CType(MainForm.Controls("chkDorsal"), Object).Value = 1 Then
			nDorsalOffset = 0
			Reinforced.n = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!chkSlantedInserts.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If CType(MainForm.Controls("chkSlantedInserts"), Object).Value = 1 Then nDorsalOffset = nInsert
			For ii = 1 To 5
				If g_Labels(ii).Found = True Then
					Reinforced.n = Reinforced.n + 1
					ARMDIA1.PR_MakeXY(xyPt, (g_Labels(ii).X / ISCALE) + 0.25 + nDorsalOffset, g_Labels(ii).y / ISCALE)
					PR_RevisePoint(xyPt)
					Reinforced.X(Reinforced.n) = xyPt.X
					Reinforced.y(Reinforced.n) = xyPt.y
				End If
			Next ii
			
			'Draw reinforcing lines using Long Dash line type
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "SetStyle" & QC & "Table(" & QQ & "find" & QCQ & "style" & QCQ & "Long Dash" & QQ & "));")
			
			ARMDIA1.PR_DrawFitted(Reinforced)
			'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!chkPalm.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If CType(MainForm.Controls("chkPalm"), Object).Value <> 1 Then
				'Draw wrist line only if not drawn above
				ARMDIA1.PR_MakeXY(xyStart, g_Labels(10).X / ISCALE, g_Labels(10).y / ISCALE)
				ARMDIA1.PR_MakeXY(xyEnd, g_Labels(11).X / ISCALE, g_Labels(11).y / ISCALE)
				PR_RevisePoint(xyStart)
				PR_RevisePoint(xyEnd)
				ARMDIA1.PR_DrawLine(xyStart, xyEnd)
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "SetStyle" & QC & "Table(" & QQ & "find" & QCQ & "style" & QCQ & "bylayer" & QQ & "));")
			
			'Add REINFORCED text
			TmpX = (g_Labels(5).X + ((g_Labels(11).X - g_Labels(5).X) / 2)) / ISCALE
			TmpY = g_Labels(11).y / ISCALE
			'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!chkPalm.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If CType(MainForm.Controls("chkPalm"), Object).Value = 1 Then
				'Shift text so that it misses PALM Text
				If CType(MainForm.Controls("txtSide"), Object).Text = "Right" Then
					TmpY = TmpY - 0.4
				Else
					TmpY = TmpY + 0.4
				End If
			End If
			ARMDIA1.PR_MakeXY(xyInsert, TmpX, TmpY)
			PR_RevisePoint(xyInsert)
			ARMDIA1.PR_DrawText("REINFORCED\nDORSAL", xyInsert, 0.125)
			
		End If
		
		'Missing Fingers
		'Check for Fingers tip options
		Static Closed(4) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optLittleTip().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Closed(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Closed(0) = CType(MainForm.Controls("optLittleTip"), Object)(0).Value
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optRingTip().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Closed(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Closed(1) = CType(MainForm.Controls("optRingTip"), Object)(0).Value
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optMiddleTip().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Closed(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Closed(2) = CType(MainForm.Controls("optMiddleTip"), Object)(0).Value
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optIndexTip().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Closed(3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Closed(3) = CType(MainForm.Controls("optIndexTip"), Object)(0).Value
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!optThumbTip().Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Closed(4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Closed(4) = CType(MainForm.Controls("optThumbTip"), Object)(0).Value
		
		nXStep = ((g_Labels(1).X - g_Labels(5).X) / 4) / ISCALE
		nYStep = ((g_Labels(1).y - g_Labels(5).y) / 4) / ISCALE
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextAngle = 0
		ARMDIA1.PR_MakeXY(xyStart, (g_Labels(5).X / ISCALE) + (nXStep / 2), (g_Labels(5).y / ISCALE) + (nYStep / 2))
		PR_RevisePoint(xyStart)
		
		'Loop through the fingers
		For ii = 0 To 3
			'        xyInsert.x = xyStart.x + (nXStep * ii) + .25
			'        xyInsert.y = xyStart.y + (nYStep * ii)
			'       Because I know about the translation and rotation
			xyInsert.X = xyStart.X + (nYStep * ii)
			xyInsert.y = xyStart.y + (nXStep * ii) + 0.25
			
			If Val(CType(MainForm.Controls("txtLen"), Object)(ii).Text) = 0 Then
				'Missing Finger
				'UPGRADE_WARNING: Couldn't resolve default property of object Closed(ii). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Closed(ii) = True Then
					ARMDIA1.PR_DrawText("CLOSED", xyInsert, 0.125)
				Else
					ARMDIA1.PR_DrawText("OPEN", xyInsert, 0.125)
				End If
			End If
			
		Next ii
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextAngle = 0
		
		'Place markers at lowest Web positions
		'For latter use with Zips, if a slant insert given then
		'supply the length of the slant insert
		ARMDIA1.PR_SetLayer("Construct")
		
		'Palmer and LFS zipper
		'based on little finger web
		If g_Labels(4).Found = True Then
			'Use little finger web
			ARMDIA1.PR_MakeXY(xyInsert, g_Labels(4).X / ISCALE, g_Labels(4).y / ISCALE)
			PR_RevisePoint(xyInsert)
		Else
			'Use start of Little Finger Side Arc
			ARMDIA1.PR_MakeXY(xyInsert, g_Labels(5).X / ISCALE, g_Labels(5).y / ISCALE)
			PR_RevisePoint(xyInsert)
		End If
		ARMDIA1.PR_MakeXY(xyPt, xyInsert.X, xyInsert.y)
		BDYUTILS.PR_DrawMarker(xyPt)
		BDYUTILS.PR_AddDBValueToLast("Zipper", "PALMER-WEB")
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!chkSlantedInserts.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CType(MainForm.Controls("chkSlantedInserts"), Object).Value = 1 Then
			BDYUTILS.PR_AddDBValueToLast("ZipperLength", Str(nInsert))
		Else
			BDYUTILS.PR_AddDBValueToLast("ZipperLength", "0")
		End If
		
		'Get diff between Little finger side and thumb side
		'to be used if a finger is missing
		nXStep = (g_Labels(1).X - g_Labels(5).X) / ISCALE
		nYStep = (g_Labels(1).y - g_Labels(5).y) / ISCALE
		
		'Dorsal
		If g_Labels(3).Found = True Then
			ARMDIA1.PR_MakeXY(xyInsert, (g_Labels(3).X / ISCALE), g_Labels(3).y / ISCALE)
			PR_RevisePoint(xyInsert)
		Else
			'Use mid point of line joining Little finger side and thumb side
			ARMDIA1.PR_MakeXY(xyInsert, (g_Labels(5).X / ISCALE) + (nXStep / 2), (g_Labels(5).y / ISCALE) + (nYStep / 2))
			PR_RevisePoint(xyInsert)
		End If
		ARMDIA1.PR_MakeXY(xyPt, xyInsert.X, xyInsert.y)
		BDYUTILS.PR_DrawMarker(xyPt)
		BDYUTILS.PR_AddDBValueToLast("Zipper", "DORSAL-WEB")
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!chkSlantedInserts.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CType(MainForm.Controls("chkSlantedInserts"), Object).Value = 1 Then
			BDYUTILS.PR_AddDBValueToLast("ZipperLength", Str(nInsert))
		Else
			BDYUTILS.PR_AddDBValueToLast("ZipperLength", "0")
		End If
		
		'Outside (Thumb Side)
		If g_Labels(2).Found = True Then
			ARMDIA1.PR_MakeXY(xyInsert, g_Labels(2).X / ISCALE, g_Labels(2).y / ISCALE)
			PR_RevisePoint(xyInsert)
		Else
			'Use end of thumb side
			ARMDIA1.PR_MakeXY(xyInsert, g_Labels(1).X / ISCALE, g_Labels(1).y / ISCALE)
			PR_RevisePoint(xyInsert)
		End If
		ARMDIA1.PR_MakeXY(xyPt, xyInsert.X, xyInsert.y)
		BDYUTILS.PR_DrawMarker(xyPt)
		BDYUTILS.PR_AddDBValueToLast("Zipper", "OUTSIDE-WEB")
		'UPGRADE_WARNING: Couldn't resolve default property of object MainForm!chkSlantedInserts.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CType(MainForm.Controls("chkSlantedInserts"), Object).Value = 1 Then
			BDYUTILS.PR_AddDBValueToLast("ZipperLength", Str(nInsert))
		Else
			BDYUTILS.PR_AddDBValueToLast("ZipperLength", "0")
		End If
		
		
		'Wrist points for extension
		ARMDIA1.PR_MakeXY(xyPalm(6), g_Labels(19).X / ISCALE, g_Labels(19).y / ISCALE)
		PR_RevisePoint(xyPalm(6))
		
		ARMDIA1.PR_MakeXY(xyPalm(5), g_Labels(22).X / ISCALE, g_Labels(22).y / ISCALE)
		PR_RevisePoint(xyPalm(5))
		
		ARMDIA1.PR_MakeXY(xyPalm(4), g_Labels(9).X / ISCALE, g_Labels(9).y / ISCALE)
		PR_RevisePoint(xyPalm(4))
		
		ARMDIA1.PR_MakeXY(xyPalm(3), g_Labels(5).X / ISCALE, g_Labels(5).y / ISCALE)
		PR_RevisePoint(xyPalm(3))
		
		ARMDIA1.PR_MakeXY(xyPalm(1), g_Labels(11).X / ISCALE, g_Labels(11).y / ISCALE)
		PR_RevisePoint(xyPalm(1))
		
		BDYUTILS.PR_DrawMarker(xyPalm(1))
		BDYUTILS.PR_AddDBValueToLast("Zipper", "PALMER")
		ARMDIA1.PR_NamedHandle("hPalmer")
		
		For ii = 3 To 6
			BDYUTILS.PR_DrawMarker(xyPalm(ii))
			BDYUTILS.PR_AddDBValueToLast("Data", "PALM" & Trim(Str(ii)))
		Next ii
		
	End Sub
	
	
	
	Sub PR_DrawAA(ByRef xCenter As Double, ByRef Ycenter As Double, ByRef nSweepAngle As Double)
		'Draw the Arc based on the above
		'Absolute co-ordinates
		'
		'INPUT
		'
		'  S         Stack of points. (Global)
		'            The Top point is always the current point
		'
		'  Xcenter#  Center of Arc in absolute co-ordinates
		'  Ycenter#     "    "   "   "   "   "   "   "
		'            in 1000th of an Inch
		'
		'  SweepAngle
		'            Angle in degrees through which the arc is
		'            swept
		'            -ve Clockwise
		'            +ve Anti-Clockwise
		'
		'OUTPUT
		'  Drafix Macro command to draw an arc
		'
		Dim Xstart, Ystart As Double
		Dim xyCenter, xyStart, xyLast As XY
		Dim nEndAngle, nStartAngle, nRadius As Double
		
		
		'Get start point from stack
		StackCurrent(Xstart, Ystart)
		ARMDIA1.PR_MakeXY(xyStart, Xstart / ISCALE, Ystart / ISCALE)
		PR_RevisePoint(xyStart)
		
		'Make center point
		ARMDIA1.PR_MakeXY(xyCenter, xCenter / ISCALE, Ycenter / ISCALE)
		PR_RevisePoint(xyCenter)
		
		'Get the angle
		nStartAngle = ARMDIA1.FN_CalcAngle(xyCenter, xyStart)
		
		'Calculate the position of the last point
		nRadius = ARMDIA1.FN_CalcLength(xyCenter, xyStart)
		
		'Draw the arc
		'Assumes that FN_Open has been called first
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hEnt = AddEntity(" & QQ & "arc" & QC & "xyStart.x +" & Str(xyCenter.X) & CC & "xyStart.y +" & Str(xyCenter.y) & CC & Str(nRadius) & CC & Str(nStartAngle) & CC & Str(nSweepAngle) & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData(hEnt," & QQ & "ID" & QQ & ",sID);")
		
		
	End Sub
	
	Sub PR_DrawEOS()
		'Draw the points stored in the stack as a line.
		'This is a special variant of the procedure PR_DrawStack
		'
		'INPUT
		'
		'  S         Stack of points. (Global)
		'            The Top point is always the current point
		'
		'OUTPUT
		'  S         The stack will contain one point
		'            this will be the current point
		'
		'  DRAFIX Macro
		'            By using a stack the line is actually drawn
		'            in reverse to the input.
		'
		'NOTES
		'    If stack is empty then or contains only one point
		'    then nothing is drawn
		'
		
		Dim X, Xprev, Xlast, Ylast, Yprev, y As Double
		Dim xyPt1, xyPt2 As XY
		
		'Return if empty
		If EmptyStack() Then Exit Sub
		
		'Store last point
		PopStack(Xlast, Ylast)
		
		'Check if empty (=> only one point)
		If EmptyStack() Then
			PushStack(Xlast, Ylast) 'Restore stack to original state
			Exit Sub
		End If
		
		'Restore stack to original state
		ARMDIA1.PR_MakeXY(xyPt1, Xlast / ISCALE, Ylast / ISCALE)
		PR_RevisePoint(xyPt1)
		
		PopStack(X, y)
		ARMDIA1.PR_MakeXY(xyPt2, X / ISCALE, y / ISCALE)
		PR_RevisePoint(xyPt2)
		
		ClearStack()
		
		'Draw line in reverse
		ARMDIA1.PR_DrawLine(xyPt2, xyPt1)
		
		'Stack will now contain only one point
		PushStack(Xlast, Ylast)
		
	End Sub
	Public Sub PR_DrawGloveHPGL_File(ByRef sFileName As String)
		'Using the HPGL data file created by the Glove Basic
		'programme create a macro to draw the data
		'
		'Modules;
		'    miFileHPGL
		'NOTE
		'    This is not a complete HPGL interpreter. It is
		'    designed only to draw the sub-set of HPGL commands
		'    used by version 3.0 of the CAD Glove Basic Programme
		'
		Dim TmpX, TmpY As Double
		Dim ii As Short
		Dim PU As Short 'Boolean Pen Up
		Dim PD As Short 'Boolean Pen Down
		Dim xyArcCen As XY
		Dim aArcSweep As Double
		
		
		'Values from FN_GetCommand
		Dim sMNemomic As String
		Dim nParams As Short
		
		'Initialise
		ClearStack()
		
		miFileHPGL = FreeFile
		FileOpen(miFileHPGL, sFileName, OpenMode.Input)
		
		'Intialise
		PU = True
		PD = False
		
		While FN_GetCommand(sMNemomic, nParams, Value)
			Select Case sMNemomic
				Case "IN"
					'Ignore
					
				Case "IP"
					'Ignore
					
				Case "PU"
					If PD Then
						PR_DrawStack()
					End If
					PU = True
					PD = False
					If nParams > 0 Then
						ClearStack()
						'UPGRADE_WARNING: Couldn't resolve default property of object Value(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						PushStack(Val(Value(nParams - 1)), Val(Value(nParams)))
					End If
					
				Case "PD"
					PU = False
					PD = True
					If nParams > 0 Then
						For ii = 1 To nParams Step 2
							'UPGRADE_WARNING: Couldn't resolve default property of object Value(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							PushStack(Val(Value(ii)), Val(Value(ii + 1)))
						Next ii
					End If
					
				Case "AA"
					'UPGRADE_WARNING: Couldn't resolve default property of object Value(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PR_EndPointAA(Val(Value(1)), Val(Value(2)), Val(Value(3)), TmpX, TmpY)
					If PD Then
						PR_DrawStack()
						'UPGRADE_WARNING: Couldn't resolve default property of object Value(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						PR_DrawAA(Val(Value(1)), Val(Value(2)), Val(Value(3)))
						'Keep these values to allow the LFA arc to be stored
						'UPGRADE_WARNING: Couldn't resolve default property of object Value(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						xyArcCen.X = Val(Value(1))
						'UPGRADE_WARNING: Couldn't resolve default property of object Value(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						xyArcCen.y = Val(Value(2))
						'UPGRADE_WARNING: Couldn't resolve default property of object Value(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						aArcSweep = Val(Value(3))
					End If
					ClearStack()
					PushStack(TmpX, TmpY)
					
				Case "PA"
					If nParams > 0 Then
						If PD Then
							If nParams > 0 Then
								For ii = 1 To nParams Step 2
									'UPGRADE_WARNING: Couldn't resolve default property of object Value(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									PushStack(Val(Value(ii)), Val(Value(ii + 1)))
								Next ii
							End If
						Else
							ClearStack()
							'UPGRADE_WARNING: Couldn't resolve default property of object Value(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							PushStack(Val(Value(nParams - 1)), Val(Value(nParams)))
						End If
					End If
					
				Case "SI"
					'Ignore
					
				Case "DI"
					'Ignore
					
				Case "LB"
					'This is used to pass info to the program.
					'The actual LB text value is not used except for
					'ii = 9 where this is the INSERT value
					'
					'Enable for De-Bug only
					'DrawText Value(1)
					
					Select Case Value(1)
						Case "EOS1", "EOS2", "EOS3"
							'End of Support Line
							'EOS1 => 2 Tapes beyond wrist
							'EOS2 => 1 Tape beyond wrist
							'EOS3 => 0 Tapes beyond wrist
							If PD Then
								If g_ExtendTo <> GLOVE_NORMAL Then
									ARMDIA1.PR_SetLayer("Construct")
									PR_DrawStack()
									ARMDIA1.PR_SetLayer("Template" & g_sSide)
								Else
									'This is a very special case as
									'the direction of the EOS line is
									'very important to the zipper modules.
									'We complicate here to simplify later (OK!)
									PR_DrawEOS()
									BDYUTILS.PR_AddDBValueToLast("Zipper", "EOS")
								End If
							End If
							
						Case "LFS"
							'End of Little Finger Side line
							If PD Then
								If g_ExtendTo <> GLOVE_NORMAL Then
									ARMDIA1.PR_SetLayer("Construct")
									PR_DrawStack()
									ARMDIA1.PR_SetLayer("Template" & g_sSide)
								Else
									PR_DrawStack()
									BDYUTILS.PR_AddDBValueToLast("Zipper", "LFS")
									BDYUTILS.PR_AddDBValueToLast("ZipperLength", "%Length")
								End If
							End If
							
						Case "ETS"
							'End of Thumb Side line
							If PD Then
								If g_ExtendTo <> GLOVE_NORMAL Then
									ARMDIA1.PR_SetLayer("Construct")
									PR_DrawStack()
									ARMDIA1.PR_SetLayer("Template" & g_sSide)
								Else
									PR_DrawStack()
									BDYUTILS.PR_AddDBValueToLast("Zipper", "ETS")
									BDYUTILS.PR_AddDBValueToLast("ZipperLength", "%Length")
								End If
							End If
							
						Case "LFA"
							'End of Little Finger Side Arc
							If g_ExtendTo <> GLOVE_NORMAL Then
								'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
								CADUTILS.PR_SetEntityData("layer", "Construct", System.DBNull.Value, System.DBNull.Value)
								BDYUTILS.PR_AddDBValueToLast("curvetype", "LFA-CONSTRUCT")
							Else
								BDYUTILS.PR_AddDBValueToLast("Zipper", "LFA")
								BDYUTILS.PR_AddDBValueToLast("ZipperLength", "%Length")
							End If
							xyLFAcen.X = xyArcCen.X / ISCALE
							xyLFAcen.y = xyArcCen.y / ISCALE
							PR_RevisePoint(xyLFAcen)
							aLFAsweep = aArcSweep
							ARMDIA1.PR_SetLayer("Template" & g_sSide)
							
						Case "MTA"
							'Middle of Thumb Arc
							ARMDIA1.PR_SetLayer("Construct")
							DrawMarker()
							BDYUTILS.PR_AddDBValueToLast("Zipper", "DORSAL")
							ARMDIA1.PR_SetLayer("Notes") 'Bit of a bodge here
							
						Case Else
							'Numeric values stored and used in PR_AddNotesEtc
							'UPGRADE_WARNING: Couldn't resolve default property of object Value(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							ii = Val(Mid(Value(1), 1, 2))
							g_Labels(ii).Found = True
							'UPGRADE_WARNING: Couldn't resolve default property of object Value(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							g_Labels(ii).sString = Value(1)
							StackCurrent(g_Labels(ii).X, g_Labels(ii).y)
							
					End Select
					
				Case "LT"
					If PD Then
						PR_DrawStack()
					End If
					SetLineType(Value(1))
					
				Case Else
					'Ignore
					
			End Select
			
		End While
		
		FileClose(miFileHPGL)
		
		'Draw last bit
		If Not PU Then
			PR_DrawStack()
		End If
		
		'Add Labels etc
		PR_AddNotesEtc()
		
		'    Close #fNum
		
	End Sub
	
	Sub PR_DrawStack()
		'Draw the point stored in the stack as a Polyline
		'
		'INPUT
		'
		'  S         Stack of points. (Global)
		'            The Top point is always the current point
		'
		'OUTPUT
		'  S         The stack will contain one point
		'            this will be the current point
		'
		'  DRAFIX Macro
		'            By using a stack the line is actually drawn
		'            in reverse to the input.
		'
		'NOTES
		'    If stack is empty then or contains only one point
		'    then nothing is drawn
		'
		
		Dim X, Xprev, Xlast, Ylast, Yprev, y As Double
		Dim xyPt As XY
		
		'UPGRADE_WARNING: Arrays in structure PolyLine may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim PolyLine As Curve
		
		'Return if empty
		If EmptyStack() Then Exit Sub
		
		'Store last point
		PopStack(Xlast, Ylast)
		
		'Check if empty (=> only one point)
		If EmptyStack() Then
			PushStack(Xlast, Ylast) 'Restore stack to original state
			Exit Sub
		End If
		
		'Restore stack to original state
		PushStack(Xlast, Ylast)
		
		'Impossible values for testing (I hope)
		PolyLine.n = 0 'Curve is set to empty
		
		Do While Not EmptyStack()
			PopStack(X, y)
			'Add point to curve
			PolyLine.n = PolyLine.n + 1
			ARMDIA1.PR_MakeXY(xyPt, X / ISCALE, y / ISCALE)
			PR_RevisePoint(xyPt)
			PolyLine.X(PolyLine.n) = xyPt.X
			PolyLine.y(PolyLine.n) = xyPt.y
		Loop 
		
		'Draw only if more than one point
		'We don't need to explicty check this as it is part of PR_DrawPoly
		ARMDIA1.PR_DrawPoly(PolyLine)
		
		'Stack will now contain only one point
		PushStack(Xlast, Ylast)
		
	End Sub
	
	Sub PR_EndPointAA(ByRef xCenter As Double, ByRef Ycenter As Double, ByRef SweepAngle As Double, ByRef Xlast As Double, ByRef Ylast As Double)
		'Returns in the variables Xlast and Ylast the value of
		'the endpoint of the arc.
		'Absolute co-ordinates
		'
		'INPUT
		'
		'  S         Stack of points. (Global)
		'            The Top point is always the current point
		'
		'  Xcenter#  Center of Arc in absolute co-ordinates
		'  Ycenter#     "    "   "   "   "   "   "   "
		'
		'  SweepAngle
		'            Angle in degrees through which the arc is
		'            swept
		'            -ve Clockwise
		'            +ve Anti-Clockwise
		'
		'OUTPUT
		'  Xlast     X and Y co-ordinate of the end point of the
		'  Ylast     arc.  The start point being the current
		'            point given by the Stack S
		'
		Dim Xstart, Ystart As Double
		Dim xyCenter, xyStart, xyLast As XY
		Dim nAngle, nRadius As Double
		
		'Get start point from stack
		StackCurrent(Xstart, Ystart)
		ARMDIA1.PR_MakeXY(xyStart, Xstart, Ystart)
		
		'Make center point
		ARMDIA1.PR_MakeXY(xyCenter, xCenter, Ycenter)
		
		'Get the angle
		nAngle = ARMDIA1.FN_CalcAngle(xyCenter, xyStart)
		
		'Revise angle wrt to sweep angle
		nAngle = nAngle + SweepAngle
		If nAngle > 360 Then
			nAngle = nAngle - 360
		ElseIf nAngle < 0 Then 
			nAngle = nAngle + 360
		End If
		
		'Calculate the position of the last point
		nRadius = ARMDIA1.FN_CalcLength(xyCenter, xyStart)
		ARMDIA1.PR_CalcPolar(xyCenter, nAngle, nRadius, xyLast)
		
		'Return values
		Xlast = xyLast.X
		Ylast = xyLast.y
		
	End Sub
	
	Sub PR_RevisePoint(ByRef xyPt As XY)
		'Procedure to translate and rotate a point so that
		'the final drawing of the glove is vertical rather than
		'the default horizontal
		'
		'
		'Constants in use
		'  XTRANSLATE
		'  YTRANSLATE
		'  ROTATION
		'  xyOrigin
		
		xyPt.X = xyPt.X + XTRANSLATE
		ARMDIA1.PR_CalcPolar(xyOrigin, ARMDIA1.FN_CalcAngle(xyOrigin, xyPt) + ROTATION, ARMDIA1.FN_CalcLength(xyOrigin, xyPt), xyPt)
		xyPt.y = xyPt.y + YTRANSLATE
		
	End Sub
	
	
	Sub PushStack(ByVal X As Double, ByVal y As Double)
		
		If S.Top = MAXSTACK Then
			MsgBox("Stack is Full", 16, "Stack Operations")
			End
		Else
			S.Top = S.Top + 1
			S.X(S.Top) = X
			S.y(S.Top) = y
			
		End If
		
	End Sub
	
	Sub SetLineType(ByRef Style As Object)
		
		Select Case Style
			Case 5
				'Line on layer notes
				ARMDIA1.PR_SetLayer("Notes")
			Case Else
				'Default to solid line
				ARMDIA1.PR_SetLayer("Template" & g_sSide)
		End Select
		
	End Sub
	
	Sub StackCurrent(ByRef X As Double, ByRef y As Double)
		'Retrieves the current position on the stack
		If Not EmptyStack() Then
			X = S.X(S.Top)
			y = S.y(S.Top)
		Else
			'Default to 0, Not particularly usefull
			'but saves an error message
			X = 0
			y = 0
		End If
	End Sub
	
	Public Function FN_GetCommand(ByRef Mnemonic As String, ByRef nParameter As Short, ByRef Parameter() As Object) As Boolean
		'Purpose:    Return an next HPGL command and it's parameters
		'Inputs:     miFileHPGL  file number of the sequentially opened HPGL file.
		'
		'Returns     True if a command found
		'            False if EOF HPGL file of HPGL file not openned
		'Arguments:  Mnemonic   - The HPGL Command string
		'            nParameter - Number of parameters for the command
		'            Parameter  - An array containg the command parameters
		'
		'Notes:      This procedure is designed to process only the restricted
		'            number of HPGL commands used by Kevin.  It is not a general
		'            tool!
		'
		'Known Bugs:
		'            Doesn't handle LB command very well (see notes)
		'
		
		'Scan the input stream breaking it into
		'Command Mnemonic and the Arguments to that command
		'
		Static CharPrev As String
		
		'UPGRADE_NOTE: Char was upgraded to Char_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Char_Renamed, sParameter As String
		Dim nAsc, nParamLen As Short
		
		'Initialise
		Mnemonic = ""
		sParameter = ""
		nParamLen = 0
		nParameter = 0
		
		'Scan input stream.
		'First 2 characters will form the Mnemonic
		'get the rest of the stream up to the command terminator
		'the difficulty is that the explict terminator is ";"
		'however it can also be terminated by a new Mnemonic
		'
		'Thus
		'    PD;PA 1000 1000;PA 2000 2000 ...
		'
		'is equivelent to
		'
		'    PDPA1000 1000PA 2000 2000 ...
		'
		'The other problem is that the LB command can be followed
		'by any ASCI character and is terminated only by the ETX
		'Chr$(3)
		'
		
		FN_GetCommand = True
		
		'Exit if file not open
		If miFileHPGL < 0 Then
			FN_GetCommand = False
			Exit Function
		End If
		
		Do While Not EOF(miFileHPGL)
			If CharPrev = "" Then
				Char_Renamed = InputString(miFileHPGL, 1)
			Else
				Char_Renamed = CharPrev
				CharPrev = ""
			End If
			nAsc = Asc(Char_Renamed)
			Select Case nAsc
				Case 65 To 90, 97 To 122
					'A-Z, a-z
					If Len(Mnemonic) = 2 Then
						If Mnemonic = "LB" Then
							'Only a single parameter for "LB"
							nParameter = 1
							sParameter = sParameter & Char_Renamed
							nParamLen = nParamLen + 1
						Else
							'Store this for later use
							CharPrev = Char_Renamed
							'Save Prameter value (if any)
							If nParamLen <> 0 Then
								nParameter = nParameter + 1
								'UPGRADE_WARNING: Couldn't resolve default property of object Parameter(nParameter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Parameter(nParameter) = sParameter
							End If
							Exit Do
						End If
					Else
						Mnemonic = Mnemonic & Char_Renamed
					End If
					
				Case 59
					';
					If Len(Mnemonic) = 2 Then
						If nParamLen <> 0 Then
							nParameter = nParameter + 1
							'UPGRADE_WARNING: Couldn't resolve default property of object Parameter(nParameter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Parameter(nParameter) = sParameter
						End If
						Exit Do
					End If
					
				Case 3
					'ETX end of text, Indicates end of LB command
					nParameter = 1
					'UPGRADE_WARNING: Couldn't resolve default property of object Parameter(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Parameter(1) = sParameter
					Exit Do
					
				Case Else
					If Mnemonic <> "LB" Then
						If (nAsc = 32) Or (nAsc = 44) Then
							'Space or Comma as parameter seperators
							'Ignore leading space or empty parameter lists
							If nParamLen <> 0 Then
								nParameter = nParameter + 1
								'UPGRADE_WARNING: Couldn't resolve default property of object Parameter(nParameter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Parameter(nParameter) = sParameter
								sParameter = ""
								nParamLen = 0
							End If
						ElseIf (nAsc >= 48 And nAsc <= 57) Or nAsc = 46 Or nAsc = 45 Or nAsc = 43 Then 
							'Only 0-9 and .
							sParameter = sParameter & Char_Renamed
							nParamLen = nParamLen + 1
						End If
					Else
						'All characters for LB Mnemonic
						sParameter = sParameter & Char_Renamed
						nParamLen = nParamLen + 1
					End If
					
			End Select
			
		Loop 
		
		If EOF(miFileHPGL) Then FN_GetCommand = False
		
		
	End Function
End Module
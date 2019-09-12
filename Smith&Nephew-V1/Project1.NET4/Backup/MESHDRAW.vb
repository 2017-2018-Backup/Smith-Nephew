Option Strict Off
Option Explicit On
Friend Class meshdraw
	Inherits System.Windows.Forms.Form
	'Module:    MESHDRAW.MAK
	'Purpose:   Mesh axilla drawing
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
	'Dec 98     GG      Ported to VB5
	'
	'Notes:-
	'
	' This module is designed to be retrofitted into the
	' existing Vest code and as such it is stand alone
	' with data being transfered by file
	'
	'
	'
	
	
	Private Sub meshdraw_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		Dim sFileName As String
		Dim sTask As String
		Dim fFileNum As Short
		Dim iError As Short
		
		Dim xyMeshStart As XY
		Dim xyTmp As XY
		Dim xyProfileStart As XY
		Dim xySeamStart As XY
		Dim xyRaglanStart As XY
		Dim xyRaglanEnd As XY
		Dim nMeshLength As Double
		Dim nDistanceAlongRaglan As Double
		
		Dim MeshProfile As BiArc
		Dim MeshSeam As BiArc
		
		Dim nLength As Double
		Dim aAngle As Double
		
		'Setup error handling
		On Error GoTo ErrorStarting
		
		'Hide form
		Hide()
		
		'Check if a previous instance is running
		'If it is exit
		'UPGRADE_ISSUE: App property App.PrevInstance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		If App.PrevInstance Then End
		
		MainForm = Me
		
		'Start a timer
		'This ensures that the dialogue dies in event of a failure
		'on the drafix macro side
		Timer1.Interval = 6000 'Approx 6 Seconds
		Timer1.Enabled = True
		
		'Find the path to the jobst system
		g_sPathJOBST = fnPathJOBST()
		
		'Open Data file
		'Note that this is a fixed file name
		sFileName = "C:\JOBST\MESHDRAW.DAT"
		fFileNum = FreeFile
		If FileLen(sFileName) = 0 Then
			MsgBox("Can't open the file " & sFileName & ".  Unable to draw the mesh.", 48)
			End
		End If
		
		'Extract from BR_MESH.D
		'hCurve = Open ("file", "C:\\JOBST\\MESHDRAW.DAT", "write") ;
		' PrintFile(hCurve, xyAxillaConstruct_2, "\n") ;
		' PrintFile(hCurve, xyAxilla, "\n") ;
		' PrintFile(hCurve, xyAxillaBodySuit, "\n") ;
		' PrintFile(hCurve, xyBackNeck, "\n") ;
		' PrintFile(hCurve, nMeshLength, "\n") ;
		' PrintFile(hCurve, nDistanceAlongRaglan, "\n") ;
		' PrintFile(hCurve, sSleeve, "\n") ;
		' PrintFile(hCurve, sID, "\n") ;
		'Close ("file", hCurve) ;
		
		
		FileOpen(fFileNum, sFileName, OpenMode.Input, OpenAccess.Read)
		Input(fFileNum, xyMeshStart.X)
		Input(fFileNum, xyMeshStart.y)
		Input(fFileNum, xyAxilla.X)
		Input(fFileNum, xyAxilla.y)
		Input(fFileNum, xyRaglanStart.X)
		Input(fFileNum, xyRaglanStart.y)
		Input(fFileNum, xyRaglanEnd.X)
		Input(fFileNum, xyRaglanEnd.y)
		Input(fFileNum, nMeshLength)
		Input(fFileNum, nDistanceAlongRaglan)
		Input(fFileNum, g_sSide)
		Input(fFileNum, g_sID)
		FileClose(fFileNum)
		
		'Offset xyMeshStart 1/8" from line (xyMeshStart,xyAxilla)
		'to allow for seam
		aAngle = ARMDIA1.FN_CalCAngle(xyMeshStart, xyAxilla) - 90
		ARMDIA1.PR_CalcPolar(xyMeshStart, aAngle, 0.125, xyProfileStart)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = ARMDIA1.FN_Open("C:\JOBST\DRAW_1.D")
		ARMDIA1.PR_SetLayer("Construct")
		BDYUTILS.PR_DrawMarker(xyProfileStart)
		
		'Body suit
		iError = FN_CalcAxillaMesh(xyRaglanStart, xyRaglanEnd, xyProfileStart, False, nMeshLength, nDistanceAlongRaglan, MeshSeam, MeshProfile)
		If iError = True Then
			'Draw the mesh
			ARMDIA1.PR_SetLayer("Template" & g_sSide)
			ARMDIA1.PR_DrawArc(MeshProfile.xyR1, MeshProfile.xyStart, MeshProfile.xyTangent)
			ARMDIA1.PR_DrawArc(MeshProfile.xyR2, MeshProfile.xyTangent, MeshProfile.xyEnd)
			iError = BDYUTILS.FN_CirLinInt(xyMeshStart, xyAxilla, MeshProfile.xyR1, MeshProfile.nR1, xyTmp)
			ARMDIA1.PR_DrawLine(MeshProfile.xyStart, xyTmp)
			
			'Draw closing line
			ARMDIA1.PR_DrawLine(xyRaglanStart, xyAxilla)
			
			ARMDIA1.PR_SetLayer("Notes")
			ARMDIA1.PR_DrawArc(MeshSeam.xyR1, MeshSeam.xyStart, MeshSeam.xyTangent)
			ARMDIA1.PR_DrawArc(MeshSeam.xyR2, MeshSeam.xyTangent, MeshSeam.xyEnd)
			ARMDIA1.PR_CalcPolar(xyMeshStart, ARMDIA1.FN_CalCAngle(xyAxilla, xyMeshStart), 3, xyTmp)
			iError = BDYUTILS.FN_CirLinInt(xyTmp, xyAxilla, MeshSeam.xyR1, MeshSeam.nR1, xyTmp)
			ARMDIA1.PR_DrawLine(MeshSeam.xyStart, xyTmp)
			ARMDIA1.PR_SetLayer("Construct")
			BDYUTILS.PR_DrawMarker(MeshSeam.xyEnd)
		Else
			MsgBox("Failure to Calculate mesh. Draw manually and contact supervisor", 48)
			BDYUTILS.PR_DrawMarker(MeshSeam.xyEnd) 'Bad practice here (using a side effect)
			'but I know that this point
			'is at the nDistanceAlongRaglan
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
		sTask = fnGetDrafixWindowTitleText()
		If sTask <> "" Then
			AppActivate(sTask)
			System.Windows.Forms.SendKeys.SendWait("@c:\jobst\draw_1.d{enter}")
		Else
			MsgBox("Can't find a Drafix Drawing to update! - Mesh Drawing Terminating.")
		End If
		
		'Terminate
		End
		
ErrorStarting: 
		
		MsgBox("Error starting - Mesh Drawing Terminating.")
		End
		
	End Sub
End Class
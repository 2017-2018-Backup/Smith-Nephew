Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class whfigbrf
	Inherits System.Windows.Forms.Form
	'Project:   WHFIGURE.MAK
	'Form:      WHFIGBRF.FRM
	'Purpose:   Restricted dialog for Waist Height Brief fabric
	'           selection
	'
	'Version:   1.01
	'Date:      16.Jan.94
	'Author:    Gary George
	'-------------------------------------------------------
	'REVISIONS:
	'Date       By      Action
	'-------------------------------------------------------
	'
	'Notes:-
	'
	'   As this is a later addition to WHFIGURE it does not
	'   use the same method as the main form to transfer data
	'   to the drawing through a DDE link.
	'
	'   It uses the method of creating a macro with all the data
	'   in place which is then executed by drafix.
	
	
	Private Sub Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Cancel.Click
		End
	End Sub
	
	Private Sub whfigbrf_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim ii As Short
		'Position to center of screen
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.
		
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default ' Change pointer to default.
		
		'Set up for fabric if fabric is given
		If g_sLastFabric <> "" Then
			If UCase(VB.Left(g_sLastFabric, 3)) = "POW" Then
				optFabric(0).Checked = 1
				optFabric_CheckedChanged(optFabric.Item(0), New System.EventArgs())
			Else
				optFabric(1).Checked = 1
				optFabric_CheckedChanged(optFabric.Item(1), New System.EventArgs())
			End If
			For ii = 0 To cboFabric.Items.Count - 1
				If g_sLastFabric = VB6.GetItemString(cboFabric, ii) Then
					cboFabric.SelectedIndex = ii
					Exit For
				End If
			Next ii
		Else
			'Use POWERNET for burns
			If g_sDiagnosis = "Burns" Then
				optFabric(0).Checked = 1
				optFabric_CheckedChanged(optFabric.Item(0), New System.EventArgs())
			Else
				optFabric(1).Checked = 1
				optFabric_CheckedChanged(optFabric.Item(1), New System.EventArgs())
			End If
		End If
	End Sub
	
	Private Sub OK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OK.Click
		Dim sFabric As String
		
		'Check if afabric has been given
		sFabric = VB6.GetItemString(cboFabric, cboFabric.SelectedIndex)
		If sFabric = "" Then
			Beep()
			Exit Sub
		End If
		
		'If a fabric has beeb given then write macro file to update drawing
		'Open the DRAFIX macro file
		'Initialise Global variables
		
		'Open file
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = FreeFile
		
		'Initialise String variables
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CC = Chr(44) 'The comma ( , )
		'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NL = Chr(10) 'The new line character
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		QQ = Chr(34) 'Double quotes ( " )
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		QCQ = QQ & CC & QQ 'Quote Comma Quote ( "," )
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		QC = QQ & CC 'Quote Comma ( ", )
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CQ = CC & QQ 'Comma Quote ( ," )
		
		'Write header information etc. to the DRAFIX macro file
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileOpen(fNum, "C:\JOBST\DRAW.D", OpenMode.Output)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//DRAFIX Figuring Macro created - " & DateString & "  " & TimeString)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//Updates waist box with BRIEF Fabric")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//by Visual Basic")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "HANDLE  hBody;")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hBody = UID (" & QQ & "find" & QC & g_iUidBody & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "if (!hBody)Exit(%cancel," & QQ & "Can't find WAIST BOX to Update" & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData( hBody, " & QQ & "Fabric" & QCQ & sFabric & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
		AppActivate(fnGetDrafixWindowTitleText())
		System.Windows.Forms.SendKeys.SendWait("@C:\JOBST\DRAW.D{enter}")
		End
	End Sub
	
	'UPGRADE_WARNING: Event optFabric.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optFabric_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optFabric.CheckedChanged
		If eventSender.Checked Then
			Dim index As Short = optFabric.GetIndex(eventSender)
			
			cboFabric.Items.Clear()
			Select Case index
				Case 0
					g_POWERNET = True
					LGLEGDIA1.PR_GetComboListFromFile(cboFabric, g_sPathJOBST & "\WHFABRIC.DAT")
				Case 1
					g_JOBSTEX = True
					cboFabric.Items.Add("53 - JOBSTEX")
					cboFabric.Items.Add("55 - JOBSTEX")
					cboFabric.Items.Add("57 - JOBSTEX")
					cboFabric.Items.Add("63 - JOBSTEX")
					cboFabric.Items.Add("65 - JOBSTEX")
					cboFabric.Items.Add("67 - JOBSTEX")
					cboFabric.Items.Add("73 - JOBSTEX")
					cboFabric.Items.Add("75 - JOBSTEX")
					cboFabric.Items.Add("77 - JOBSTEX")
					cboFabric.Items.Add("83 - JOBSTEX")
					cboFabric.Items.Add("85 - JOBSTEX")
					cboFabric.Items.Add("87 - JOBSTEX")
			End Select
			
		End If
	End Sub
End Class
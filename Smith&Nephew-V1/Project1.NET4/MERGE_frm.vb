Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class merge
	Inherits System.Windows.Forms.Form
    'Project:   Merge.MAK
    'Purpose:   To allow the user to locate a cad drawing file to
    '           merge with the current darwing
    'Version:   1.00
    'Date:      06.Feb.98
    'Author:    Gary George
    '---------------------------------------------------------------------------------------------
    'REVISIONS:
    'Date       By      Action
    '---------------------------------------------------------------------------------------------
    'Dec 98     GG      Ported to VB5
    '
    'NOTE:-
    '
    '
    '
    Public fNum As Short
    '
    Public CC As String 'The comma (,)
    Public NL As String 'The new line character
    Public TB As String 'The tab char character
    Public QQ As String 'Double quotes (")
    Public QCQ As String 'QQ & CC & QQ
    Public QC As String 'QQ & CC
    Public CQ As String 'CC & QQ

    Dim m_sDrawingToMergeFullPath As String
	Dim m_sDrawingToMergeTitleOnly As String
	
	Private Sub CAD_File_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CAD_File.Click
		'Find an existing CAD file to merge
		'
		'Path to drawings
		Dim lpBuffer As New VB6.FixedLengthString(144) 'Minimum recommended wrt GetWindowsDirectory()
		Dim nBufferSize, nSize As Short
		Dim WindowsDir As String
		
		nBufferSize = 143
		
		'Get the path to the Windows Directory to locate DRAFIX.INI
		'
		nSize = GetWindowsDirectory(lpBuffer.Value, nBufferSize)
		WindowsDir = VB.Left(lpBuffer.Value, nSize)
		
		'Get the path to the Drawing directory from
		'DRAFIX.INI
		'
		nSize = GetPrivateProfileString("Path", "PathDrawing", "C:\", lpBuffer.Value, nBufferSize, WindowsDir & "\DRAFIX.INI")
		
		CMDialog1Open.InitialDirectory = VB.Left(lpBuffer.Value, nSize)
		'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		CMDialog1Open.Filter = "Drafix CAD (*.cad)|*.cad"
		CMDialog1Open.Title = "CAD File to Merge"
		CMDialog1Open.FileName = ""
		CMDialog1Open.DefaultExt = "cad"
		CMDialog1Open.ShowDialog()
		
		If CMDialog1Open.FileName <> "" Then
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If Dir(CMDialog1Open.FileName) = "" Then
				MsgBox("CAD File not found.", 48, "Merge Drawings")
			Else
				Label1.Text = CMDialog1Open.FileName
				m_sDrawingToMergeFullPath = CMDialog1Open.FileName
				'UPGRADE_WARNING: CommonDialog property CMDialog1.FileTitle was upgraded to CMDialog1.FileName which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
				m_sDrawingToMergeTitleOnly = CMDialog1Open.FileName
			End If
		End If
	End Sub
	
	Private Sub Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Cancel.Click
        Return
    End Sub
	
	Private Function FN_EscapeSlashesInString(ByRef sAssignedString As Object) As String
		'Search through the string looking for " (double quote characater)
		'If found use \ (Backslash) to escape it
		'
		Dim ii As Short
		'UPGRADE_NOTE: Char was upgraded to Char_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Char_Renamed As String
		Dim sEscapedString As String
		
		FN_EscapeSlashesInString = ""
		
		For ii = 1 To Len(sAssignedString)
			'UPGRADE_WARNING: Couldn't resolve default property of object sAssignedString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Char_Renamed = Mid(sAssignedString, ii, 1)
			If Char_Renamed = "\" Then
				sEscapedString = sEscapedString & "\" & Char_Renamed
			Else
				sEscapedString = sEscapedString & Char_Renamed
			End If
		Next ii
		
		FN_EscapeSlashesInString = sEscapedString
		
	End Function
	
	
	'UPGRADE_ISSUE: Form event Form.LinkClose was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub Form_LinkClose()
		'Disable TIMEOUT Timer
		Timer1.Enabled = False
		
		Show()
		
	End Sub
	
	Private Sub merge_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Hide form to allow DDE to happen
		'Show form on LinkClose
		Hide()
		
		'Clear Text Boxes
		'Work Order and TXF file Name
		txtWorkOrder.Text = ""
		
		'Patient Details
		txtFileNo.Text = ""
		txtPatientName.Text = ""
		txtOrderDate.Text = ""
		txtCurrentCADFile.Text = ""
		
		'Position to center of screen
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
		
		MainForm = Me
		
		'Get the path to the JOBST system installation directory
		g_sPathJOBST = fnPathJOBST()
		
		'Ensure that we timeout after approx 6 seconds
		'The timer is disabled on link close
		Timer1.Interval = 6000
		Timer1.Enabled = False
		
	End Sub
	
	Private Sub OK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OK.Click
		'Check that with the user that data is correct
		'If OK the update drawing
		Dim sDrafixTask As String
		Dim sDrawingToMerge As String
		Dim sMessage As String
		
		'Disable to stop double click
		OK.Enabled = False
		sMessage = ""
		
		'Check for existance of file
		If m_sDrawingToMergeFullPath = "" Then GoTo ExitSub_OK_Click
		
		'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If m_sDrawingToMergeFullPath = txtCurrentCADFile.Text Then sMessage = "Selected Drawing File is the same as the current Drawing File!" & NL
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Dir(m_sDrawingToMergeFullPath) = "" Then sMessage = sMessage & "CAD Drawing File not found!" & NL
		
		'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If UCase(VB.Right(m_sDrawingToMergeFullPath, 4)) <> ".CAD" Then sMessage = sMessage & "The selected file is not a CAD Drawing File!" & NL
		
		If Len(sMessage) > 0 Then
			MsgBox(sMessage, 48, "Errors selecting file")
			GoTo ExitSub_OK_Click
		End If
		
		PR_DRAFIX_Macro("C:\JOBST\DRAW.D")
		
		sDrafixTask = fnGetDrafixWindowTitleText()
		
		'Start the required task by using the appropriate
		'drafix macro
		If sDrafixTask <> "" Then
			AppActivate(sDrafixTask)
			System.Windows.Forms.SendKeys.SendWait("@C:\JOBST\DRAW.D{enter}")
            Return
        Else
			MsgBox("Unable to find a running copy of DRAFIX!", 16, "Error - Merging Drawing")
		End If
		
ExitSub_OK_Click: 
		
		OK.Enabled = True
		
	End Sub
	
	Private Sub PR_DRAFIX_Macro(ByRef sDrafixFile As String)
		'Create a DRAFIX macro file
		'NOTES
		'
		Dim sString As String
		
		'Open file
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = FreeFile
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileOpen(fNum, sDrafixFile, OpenMode.Output)
		
		'Initialise String globals
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CC = Chr(44) 'The comma (,)
		'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NL = Chr(10) 'The new line character
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		QQ = Chr(34) 'Double quotes (")
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		QCQ = QQ & CC & QQ
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		QC = QQ & CC
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CQ = CC & QQ
		
		'Write header information etc. to the DRAFIX macro file
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//DRAFIX Macro created - " & DateString & "  " & TimeString)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//Patient    - " & txtPatientName.Text & CC & " " & txtFileNo.Text & CC)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//Work Order - " & txtWorkOrder.Text)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//by Visual Basic")
		
		'    Print #fNum, "@" & FN_EscapeSlashesInString(g_sPathJOBST) & "\\CADLINK\\MERGE.D "; QQ; FN_EscapeSlashesInString((m_sDrawingToMergeFullPath)); QQ; ";"
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "@" & g_sPathJOBST & "\CADLINK\MERGE.D " & m_sDrawingToMergeFullPath & ";")
		
		'Close the Macro File
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "// -- End of MACRO --")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
	End Sub
	
	Private Sub PR_GetDrafixWindowTitleText(ByRef sDrafixText As String, ByRef sDrafixTask As String)
		'Returns the
		'    Drafix Window Title Text in sDrafixTask$
		'
		'N.B. Warning
		'
		'    Returns last Drafix task found.
		'    This is OK for Drafix 2.1e as only a single instance
		'    of Drafix Windows CAD is allowed.
		'    Dangerous for Drafix 3.
		
		Dim sTask As String
		Dim nLength As Integer
		Dim nCurrWnd As Integer
		Dim X As Short
		Dim nDrafixText, nInstanceText As Short
		
		'Get the nWnd of the first item in the master list
		'so we can process the task list entries (top level only)
		nCurrWnd = GetWindow(MainForm.Handle.ToInt32, GW_HWNDFIRST)
		
		'Loop to locate Drafix CAD task and Requested Instance
		sDrafixTask = ""
		nDrafixText = Len(sDrafixText)
		
		While nCurrWnd <> 0
			
			'Extract details of task
			nLength = GetWindowTextLength(nCurrWnd)
			sTask = Space(nLength + 1)
			nLength = GetWindowText(nCurrWnd, sTask, nLength + 1)
			
			If VB.Left(sTask, nDrafixText) = sDrafixText Then sDrafixTask = sTask
			
			'Get next task from master list
			nCurrWnd = GetWindow(nCurrWnd, GW_HWNDNEXT)
			
			'Process Windows events
			'UPGRADE_ISSUE: DoEvents does not return a value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8D115264-E27F-4472-A684-865A00B5E826"'
			X = System.Windows.Forms.Application.DoEvents()
			
		End While
		
	End Sub
	
	Private Sub PR_ReadLine(ByVal fFile As Short, ByRef sLine As String)
		'Read a line, character at a time up to either the
		'   NL character
		'or
		'   CRNL characters
		'
		'UPGRADE_NOTE: Char was upgraded to Char_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim NewLine, CR As String
		Dim Char_Renamed As New VB6.FixedLengthString(1)
		NewLine = Chr(10)
		CR = Chr(13)
		sLine = ""
		
		Do Until EOF(fFile)
			Char_Renamed.Value = InputString(fFile, 1)
			If Char_Renamed.Value = NewLine Then Exit Do
			If Char_Renamed.Value <> CR Then sLine = sLine & Char_Renamed.Value
		Loop 
		
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
        'In case there is no Link_Close event
        'The programme will time out after approx 5 secs
        Return
    End Sub
End Class
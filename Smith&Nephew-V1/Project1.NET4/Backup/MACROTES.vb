Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class MacroTest
	Inherits System.Windows.Forms.Form
	'Project:   JOBSTART.MAK
	'Purpose:   General procedure to start JOBST CAD modules.
	'
	'           Ensures that no more than one instance of
	'           of a JOBST CAD system VB module can exist.
	'           This stops the problem of DRAFIX macros opening
	'           a DDE conversation with the wrong instance.
	'
	'
	'Version:   3.01
	'Date:      10.Jun.94
	'Author:    Gary George
	'---------------------------------------------------------------------------------------------
	'REVISIONS:
	'Date       By      Action
	'---------------------------------------------------------------------------------------------
	'Dec 98     GG      Ported to VB 5
	
	'* Windows API Functions Declarations
	'    Private Declare Function GetWindow Lib "User" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer
	'    Private Declare Function GetWindowText Lib "User" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
	'    Private Declare Function GetWindowTextLength Lib "User" (ByVal hWnd As Integer) As Integer
	'    Private Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
	'    Private Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFilename$)
	'
	'   'Constanst used by GetWindow
	'    Const GW_CHILD = 5
	'    Const GW_HWNDFIRST = 0
	'    Const GW_HWNDLAST = 1
	'    Const GW_HWNDNEXT = 2
	'    Const GW_HWNDPREV = 3
	'    Const GW_OWNER = 4
	
	Public MainForm As MacroTest
	
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		On Error Resume Next
		
		Dim sDrafixTask As String
		Dim sTask As String
		Dim sMacro As String
		Dim sCommand As String
		Dim sDrafixText As String
		Dim sTaskText As String
		Dim PathJOBST As String
		
		
		sCommand = "@C:\CADLINK\Imagea.D{enter}"
		
		sDrafixText = "Drafix CAD Professional"
		
		PR_GetDrafixandTaskWindowTitleText(sDrafixTask, sTask, sDrafixText, sTaskText)
		
		'Start the required task by using the appropriate
		'drafix macro
		If sDrafixTask <> "" Then
			AppActivate(sDrafixTask)
			System.Windows.Forms.SendKeys.SendWait(sCommand)
		Else
			MsgBox("Can't find a specific drawing to work with!", 16, "Starting CAD Module")
		End If
		
		
		
		
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		Me.Close()
		End
	End Sub
	
	Private Sub MacroTest_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		'UPGRADE_ISSUE: App property App.PrevInstance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		If App.PrevInstance Then End
		
		
		MainForm = Me
		
		Show()
		
	End Sub
	
	Private Sub PR_GetDrafixandTaskWindowTitleText(ByRef sDrafixTask As String, ByRef sInstanceTask As String, ByRef sDrafixText As String, ByRef sInstanceText As String)
		'Returns the
		'    Drafix Window Title Text
		'    Requested Instance Window Title Text
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
		Dim X As Integer
		Dim nDrafixText As Integer
		Dim nInstanceText As Integer
		
		'Get the nWnd of the first item in the master list
		'so we can process the task list entries (top level only)
		nCurrWnd = GetWindow(MainForm.Handle.ToInt32, GW_HWNDFIRST)
		
		'Loop to locate Drafix CAD task and Requested Instance
		sDrafixTask = ""
		nDrafixText = Len(sDrafixText)
		sInstanceTask = ""
		nInstanceText = Len(sInstanceText)
		
		While nCurrWnd <> 0
			
			'Extract details of task
			nLength = GetWindowTextLength(nCurrWnd)
			sTask = Space(nLength + 1)
			nLength = GetWindowText(nCurrWnd, sTask, nLength + 1)
			
			If VB.Left(sTask, nDrafixText) = sDrafixText Then sDrafixTask = sTask
			If VB.Left(sTask, nInstanceText) = sInstanceText Then sInstanceTask = sTask
			
			'Get next task from master list
			nCurrWnd = GetWindow(nCurrWnd, GW_HWNDNEXT)
			
			'Process Windows events
			'UPGRADE_ISSUE: DoEvents does not return a value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8D115264-E27F-4472-A684-865A00B5E826"'
			X = System.Windows.Forms.Application.DoEvents()
			
		End While
		
	End Sub
End Class
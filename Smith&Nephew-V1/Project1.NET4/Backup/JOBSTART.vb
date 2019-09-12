Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class jobstart
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
	
	Public MainForm As jobstart
	
	
	Private Sub jobstart_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		On Error Resume Next
		
		Dim sDrafixTask As String
		Dim sTask As String
		Dim sMacro As String
		Dim sCommand As String
		Dim sDrafixText As String
		Dim sTaskText As String
		Dim PathJOBST As String
		
		'UPGRADE_ISSUE: App property App.PrevInstance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		If App.PrevInstance Then End
		
		Select Case VB.Command()
			Case "armedit"
				sMacro = "\ARM\ARM_EDT.D{enter}"
				sTaskText = "ARM Edit -"
				
			Case "waistlegedit"
				sMacro = "\WAIST\WH_LGEDT.D{enter}"
				sTaskText = "LEG Edit -"
				
			Case "leginput L"
				sMacro = "\LEG\LG_INLEG.D Left{enter}"
				sTaskText = "LEG Details -"
				
			Case "leginput R"
				sMacro = "\LEG\LG_INLEG.D Right{enter}"
				sTaskText = "LEG Details -"
				
			Case "arminput L"
				sMacro = "\ARM\ARML.D{enter}"
				sTaskText = "ARM Details -"
				
			Case "arminput R"
				sMacro = "\ARM\ARMR.D{enter}"
				sTaskText = "ARM Details -"
				
			Case "arminput S"
				sMacro = "\ARM\ARM_SEL.D{enter}"
				sTaskText = "ARM Details -"
				
			Case "cadglove L"
				sMacro = "\GLOVECAD\GL_INL.D{enter}"
				sTaskText = "CAD Glove - Dialogue"
				
			Case "cadglove R"
				sMacro = "\GLOVECAD\GL_INR.D{enter}"
				sTaskText = "CAD Glove - Dialogue"
				
			Case "manglove L"
				sMacro = "\GLOVEMAN\GLM_INL.D{enter}"
				sTaskText = "MANUAL Glove - Dialogue"
				
			Case "manglove R"
				sMacro = "\GLOVEMAN\GLM_INR.D{enter}"
				sTaskText = "MANUAL Glove - Dialogue"
				
			Case "webspacr L"
				sMacro = "\WEBSPACR\WEB_INL.D{enter}"
				sTaskText = "Web Spacers - Draw"
				
			Case "webspacr R"
				sMacro = "\WEBSPACR\WEB_INR.D{enter}"
				sTaskText = "Web Spacers - Draw"
				
			Case "cadglove S"
				sMacro = "\GLOVECAD\GL_SEL.D{enter}"
				sTaskText = "CAD Glove - Dialogue"
				
			Case "cadglove E"
				sMacro = "\GLOVECAD\GL_EDT.D{enter}"
				sTaskText = "CAD GLOVE Edit -"
				
			Case "manglove S"
				sMacro = "\GLOVEMAN\GLM_SEL.D{enter}"
				sTaskText = "MANUAL Glove - Dialogue"
				
			Case "manglove E"
				sMacro = "\GLOVEMAN\GLM_EDT.D{enter}"
				sTaskText = "GLOVE Edit -"
				
			Case "vestarminput L"
				sMacro = "\ARM\SLVL.D{enter}"
				sTaskText = "ARM Details -"
				
			Case "vestarminput R"
				sMacro = "\ARM\SLVR.D{enter}"
				sTaskText = "ARM Details -"
				
			Case "vestarminput S"
				sMacro = "\ARM\SLV_SEL.D{enter}"
				sTaskText = "ARM Details -"
				
			Case "stamp"
				sMacro = "\STAMPS\STMPSTRT.D{enter}"
				sTaskText = "Stamps"
				
			Case "stamp wh"
				sMacro = "\STAMPS\STMPSTWH.D{enter}"
				sTaskText = "Stamps"
				
			Case "stamp am"
				sMacro = "\STAMPS\STMPSTAM.D{enter}"
				sTaskText = "Stamps"
				
			Case "stamp lg"
				sMacro = "\STAMPS\STMPSTLG.D{enter}"
				sTaskText = "Stamps"
				
			Case "waistbody"
				sMacro = "\WAIST\WH_INBOD.D{enter}"
				sTaskText = "Waist Height Body - Data"
				
			Case "waistleg L"
				sMacro = "\WAIST\WH_INLEG.D Left{enter}"
				sTaskText = "Waist Height Leg Data"
				
			Case "waistleg R"
				sMacro = "\WAIST\WH_INLEG.D Right{enter}"
				sTaskText = "Waist Height Leg Data"
				
			Case "waistfigure"
				sMacro = "\WAIST\WH_FIGUR.D{enter}"
				sTaskText = "Waist Height - Figure"
				
			Case "vest"
				sMacro = "\VEST\VST_IN.D{enter}"
				sTaskText = "VEST Body - Dialogue"
				
			Case "torsoband"
				sMacro = "\VEST\TORSO_IN.D{enter}"
				sTaskText = "TORSO Band - Dialogue"
				
			Case "scoopneck"
				sMacro = "\VEST\SCOOP_IN.D{enter}"
				sTaskText = "VEST Side Scoop Neck"
				
			Case "bodysuit"
				sMacro = "\BODY\BDY_IN.D{enter}"
				sTaskText = "Body Brief/Suit - Dialogue"
				
			Case "bodysuit S"
				sMacro = "\BODY\BDY_SEL.D{enter}"
				sTaskText = "Body Brief/Suit - Dialogue"
				
			Case "bodydraw"
				sMacro = "\BODY\BDY_DRAW.D Left{enter}"
				sTaskText = "Body Draw - Dialogue"
				
			Case "bodyarminput L"
				sMacro = "\ARM\BODL.D{enter}"
				sTaskText = "ARM Details -"
				
			Case "bodyarminput R"
				sMacro = "\ARM\BODR.D{enter}"
				sTaskText = "ARM Details -"
				
			Case "bodyarminput S"
				sMacro = "\ARM\BOD_SEL.D{enter}"
				sTaskText = "ARM Details -"
				
			Case "bodyleginput R"
				sMacro = "\BODY\BD_INLEG.D Right{enter}"
				sTaskText = "Body Suit Leg Data -"
				
			Case "bodyleginput L"
				sMacro = "\BODY\BD_INLEG.D Left{enter}"
				sTaskText = "Body Suit Leg Data -"
				
			Case "TRITONtoCAD"
				sMacro = "\CADLINK\CL_TXF.D{enter}"
				sTaskText = "Patient and Work Order Details"
				
			Case "headneck"
				sMacro = "\HEADNECK\HEADNECK.D{enter}"
				sTaskText = "Head & Neck"
				
			Case "PDandWOD"
				sMacro = "\CADLINK\CL_PDWOD.D{enter}"
				sTaskText = "Patient and Work Order Details"
				
			Case Else
				End
		End Select
		
		MainForm = Me
		
		'Create Command string
		sCommand = "@" & fnPathJOBST() & sMacro
		
		sDrafixText = "Drafix CAD Professional"
		
		PR_GetDrafixandTaskWindowTitleText(sDrafixTask, sTask, sDrafixText, sTaskText)
		
		'If an instance of the required task is already running
		'then pop it to the front
		If sTask <> "" And sDrafixTask <> "" Then
			AppActivate(sTask)
			End
		End If
		
		'Start the required task by using the appropriate
		'drafix macro
		If sDrafixTask <> "" Then
			AppActivate(sDrafixTask)
			System.Windows.Forms.SendKeys.SendWait(sCommand)
			End
		Else
			MsgBox("Can't find a specific drawing to work with!", 16, "Starting CAD Module")
			End
		End If
		
		'Finish if not able to do anything
		End
		
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
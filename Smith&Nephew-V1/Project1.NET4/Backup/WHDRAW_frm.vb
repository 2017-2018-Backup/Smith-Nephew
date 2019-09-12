Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class whdraw
	Inherits System.Windows.Forms.Form
	'Project:   WHDRAW.FRM
	'Purpose:   Drawing Arms and Vest sleeves.
	'
	'Version:   3.0
	'Date:      94/95
	'Author:    Gary George
	'
	'-------------------------------------------------------
	'REVISIONS:
	'Date       By      Action
	'-------------------------------------------------------
	'Feb.95     GG      TRITON to CADLINK
	'
	'Dec 98     GG      Ported to VB5
	'
	'
	'Note:-
	'
	'This module is used as the decision making front end to
	'all of the drawing functions for the Waist Height
	'This means that the actual DRAFIX macros can be tailored to
	'load only those Macros that are relevent to the required drawing.
	'E.G.
	'    Drawing a footless leg means that WHFTPNTS.D need not be used.
	'    Drawing of a 2nd leg with no cutout means that the WHLG2DBD.D
	'    is not required.
	'    Etc.
	'
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
	
	'MsgBox constants
	'    Const IDYES = 6
	'    Const IDNO = 7
	
	
	Private Function FN_ChapOpen(ByRef sDrafixFile As String) As Short
		'Open the DRAFIX macro file that draws the CHAP
		'Initialise Global variables
		
		'Open file
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = FreeFile
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileOpen(fNum, sDrafixFile, OpenMode.Output)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FN_ChapOpen = fNum
		
		
		'Write header information etc. to the DRAFIX macro file
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//DRAFIX CHAP Drawing Macro created - " & DateString & "  " & TimeString)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//Patient - " & g_sPatient & CC & " " & g_sFileNo & CC & " SIDE - " & g_sSide)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//by Visual Basic")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "HANDLE hEnt, hLayer;")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "XY xyStart;")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "STRING sID;")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "xyStart.x=0;")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "xyStart.y=0;")
		BDYUTILS.PR_PutStringAssign("sID", txtID)
		
		
	End Function
	
	
	Private Function FN_WHOpen(ByRef sDrafixFile As String, ByRef sName As Object, ByRef sPatientFile As Object, ByRef sLeftorRight As Object) As Short
		'Open the DRAFIX macro file
		'Initialise Global variables
		
		'Open file
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = FreeFile
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileOpen(fNum, sDrafixFile, OpenMode.Output)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FN_WHOpen = fNum
		
		'Initialise patient globals
		'UPGRADE_WARNING: Couldn't resolve default property of object sPatientFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_sFileNo = sPatientFile
		'UPGRADE_WARNING: Couldn't resolve default property of object sLeftorRight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_sSide = sLeftorRight
		'UPGRADE_WARNING: Couldn't resolve default property of object sName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_sPatient = sName
		
		'Write header information etc. to the DRAFIX macro file
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//DRAFIX Waist Height Drawing Macro created - " & DateString & "  " & TimeString)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//Patient - " & g_sPatient & CC & " " & g_sFileNo & CC & " SIDE - " & g_sSide)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//by Visual Basic")
		
		'Ensure that style and color are by layer
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "SetStyle" & QC & "Table(" & QQ & "find" & QCQ & "style" & QCQ & "bylayer" & QQ & ")) ;")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "SetColor" & QC & "Table(" & QQ & "find" & QCQ & "color" & QCQ & "bylayer" & QQ & ")) ;")
		
		'Text data
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetData(" & QQ & "TextHorzJust" & QC & g_nCurrTextHorizJust & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetData(" & QQ & "TextVertJust" & QC & g_nCurrTextVertJust & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetData(" & QQ & "TextHeight" & QC & g_nCurrTextHt & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetData(" & QQ & "TextAspect" & QC & g_nCurrTextAspect & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetData(" & QQ & "TextFont" & QC & g_nCurrTextFont & ");")
		
	End Function
	
	
	Private Function fnGetNumber(ByVal sString As String, ByRef iIndex As Short) As Double
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
				sString = VB.Left(sString, iPos - 1)
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
	
	'UPGRADE_ISSUE: Form event Form.LinkClose was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub Form_LinkClose()
		Dim nAge, ii, nn, iFabricClass As Short
		Dim iZipper, iMMHg, iValue As Short
		Dim nValue, nStretch As Double
		Dim iLegStyle As Short
		
		Timer1.Enabled = False
		
		'Units
		If txtUnits.Text = "cm" Then
			g_nUnitsFac = 10 / 25.4
		Else
			g_nUnitsFac = 1
		End If
		
		'Work order
		If txtWorkOrder.Text = "" Then
			g_sWorkOrder = "-"
		Else
			g_sWorkOrder = txtWorkOrder.Text
		End If
		
		iLegStyle = ARMEDDIA1.fnGetNumber(txtLegStyle.Text, 1)
		Select Case txtDrawType.Text
			Case "WH1LG"
				g_sDialogueID = "First Leg Drawing"
				If iLegStyle = 2 Then
					'Brief
					PR_WH_Brief1stLeg()
				ElseIf iLegStyle = 7 Or iLegStyle = 6 Or iLegStyle = 5 Then 
					'ChapStyle
					PR_WH_ChapLeg((txtDrawType.Text))
				Else
					'Panty or full leg
					PR_WH_1stLeg()
				End If
				
			Case "WH2LG"
				g_sDialogueID = "Second Leg Drawing"
				If iLegStyle = 2 Then
					'Brief
					PR_WH_Brief2ndLeg()
				ElseIf iLegStyle = 7 Or iLegStyle = 6 Or iLegStyle = 5 Then 
					'ChapStyle
					PR_WH_ChapLeg((txtDrawType.Text))
				Else
					'Panty or full leg
					PR_WH_2ndLeg()
				End If
				
			Case "WHCUT"
				g_sDialogueID = "Cut Out Drawing"
				If iLegStyle = 7 Or iLegStyle = 6 Or iLegStyle = 5 Then
					MsgBox("Can't Draw a CUT-OUT for Chap styles", 16, g_sDialogueID)
					End
				End If
				PR_WH_CutOut()
				
			Case "CHAPBODYLeft", "CHAPBODYRight"
				g_sDialogueID = "Chap Drawing"
				PR_WH_ChapBody(Mid(txtDrawType.Text, 9))
				
			Case Else
				End
		End Select
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default ' Change pointer to default.
		AppActivate(fnGetDrafixWindowTitleText())
		System.Windows.Forms.SendKeys.SendWait("@C:\JOBST\DRAW.D{enter}")
		End
		'Show
	End Sub
	
	
	'UPGRADE_ISSUE: Form event Form.LinkExecute was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub Form_LinkExecute(ByRef CmdStr As String, ByRef cancel As Short)
		If CmdStr = "Cancel" Then
			cancel = 0
			End
		End If
	End Sub
	
	Private Sub whdraw_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim ii As Short
		Hide()
		'Check if a previous instance is running
		'If it is warn user and exit
		
		'Maintain while loading DDE data
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor ' Change pointer to hourglass.
		'Reset in Form_LinkClose
		
		'Position to center of screen
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.
		
		MainForm = Me
		
		g_nUnitsFac = 1 'Default to inches
		
		g_sPathJOBST = fnPathJOBST()
		
		'Initialise String globals
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
		
		
		'Globals to reduced drafix code written to file
		g_sCurrentLayer = ""
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextHt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextHt = 0.125
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextAspect = 0.6
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextHorizJust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextHorizJust = 1 'Left
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextVertJust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextVertJust = 8 'Bottom
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextFont. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextFont = 0 'BLOCK
		
		
		'Process will Time Out after approx 10 secs
		Timer1.Interval = 10000
		Timer1.Enabled = True
		
	End Sub
	
	Private Sub OK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OK.Click
		'Not available in normal use
		Dim sTask As String
		sTask = fnGetDrafixWindowTitleText()
		AppActivate(sTask)
		System.Windows.Forms.SendKeys.SendWait("@C:\JOBST\DRAW.D{enter}")
		End
	End Sub
	
	Private Sub PR_PutLine(ByRef sLine As String)
		'Puts the contents of sLine to the opened "Macro" file
		'Puts the line with no translation or additions
		'    fNum is global variable
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, sLine)
		
	End Sub
	
	Private Sub PR_PutNumberAssign(ByRef sVariableName As String, ByRef nAssignedNumber As Object)
		
		'Procedure to put a number assignment
		'Adds a semi-colon
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nAssignedNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, sVariableName & "=" & Str(nAssignedNumber) & ";")
		
		
	End Sub
	
	Private Sub PR_PutStringAssign(ByRef sVariableName As String, ByRef sAssignedString As Object)
		'Procedure to put a string assignment
		'Encloses String in quotes and adds a semi-colon
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, sVariableName & "=" & QQ & sAssignedString & QQ & ";")
		
	End Sub
	
	Private Sub PR_StandardDetails()
		Dim nAge As Short
		
		'Path to JOBST
		BDYUTILS.PR_PutStringAssign("sPathJOBST", ARMDIA1.FN_EscapeSlashesInString(g_sPathJOBST))
		
		'Patient Details
		BDYUTILS.PR_PutStringAssign("sPatient", txtPatientName)
		BDYUTILS.PR_PutStringAssign("sFileNo", txtFileNo)
		BDYUTILS.PR_PutStringAssign("sWorkOrder", g_sWorkOrder)
		
		BDYUTILS.PR_PutStringAssign("sAge", txtAge)
		nAge = Val(txtAge.Text)
		BDYUTILS.PR_PutNumberAssign("nAge", nAge)
		
		BDYUTILS.PR_PutStringAssign("sSEX", txtSex)
		If txtSex.Text = "Male" Then
			BDYUTILS.PR_PutLine("Male = %true;")
			BDYUTILS.PR_PutLine("Female = %false;")
		Else
			BDYUTILS.PR_PutLine("Male = %false;")
			BDYUTILS.PR_PutLine("Female = %true;")
		End If
		
		BDYUTILS.PR_PutStringAssign("sUnits", txtUnits)
		BDYUTILS.PR_PutNumberAssign("nUnitsFac", g_nUnitsFac)
		
		BDYUTILS.PR_PutStringAssign("sDiagnosis", txtDiagnosis)
		
	End Sub
	
	Private Sub PR_WH_1stLeg()
		'Procedure to create a Macro to draw the first leg of a
		'Waist Height
		Dim sFile, sSide As String
		Dim sLengths As String
		Dim nLeftLegStyle, nLegStyle, nRightLegStyle As Short
		Dim nLastTape, nFirstTape, nFabricClass As Short
		Dim Footless, ii, itemplate As Short
		Dim nFoldHt, nValue As Double
		
		nLegStyle = ARMEDDIA1.fnGetNumber(txtLegStyle.Text, 1)
		nLeftLegStyle = ARMEDDIA1.fnGetNumber(txtLegStyle.Text, 2)
		nRightLegStyle = ARMEDDIA1.fnGetNumber(txtLegStyle.Text, 3)
		
		sFile = "C:\JOBST\DRAW.D"
		
		'Establish which is the First Leg
		Select Case nLegStyle
			Case 0, 1
				If ARMEDDIA1.fnDisplaytoInches(Val(txtLeftThighCir.Text)) - ARMEDDIA1.fnDisplaytoInches(Val(txtRightThighCir.Text)) > 1 Then
					sSide = "Right"
				Else
					sSide = "Left"
				End If
			Case 2
				sSide = "Left"
			Case 3, 5
				sSide = "Left"
			Case 4, 6
				sSide = "Right"
		End Select
		
		'Having decided on leg do a swift check to see if there is enough
		'data to draw the leg
		If sSide = "Left" And txtLeftAnkle.Text = "" And txtLeftLengths.Text = "" Then
			MsgBox("No data for LEFT Leg", 16, g_sDialogueID)
			End
		End If
		
		If sSide = "Left" And txtLeftTemplate.Text = "" Then
			MsgBox("Figure LEFT Leg before drawing", 16, "First Leg Drawing")
			End
		End If
		
		If sSide = "Right" And txtRightAnkle.Text = "" And txtRightLengths.Text = "" Then
			MsgBox("No data for RIGHT Leg", 16, g_sDialogueID)
			End
		End If
		
		If sSide = "Right" And txtRightTemplate.Text = "" Then
			MsgBox("Figure RIGHT Leg before drawing", 16, g_sDialogueID)
			End
		End If
		
		
		'    If sSide = "Left" And fnGetNumber(txtLeftAnkle.Text, 1) = -1 And nLeftLegStyle <> 1 Then
		'        MsgBox "A Full leg has been chosen in the Body, but the data in the leg is footless.  ", 16, "First Leg Drawing"
		'        End
		'    End If
		
		
		'Open Macro file (fNum is declared as Global)
		sFile = "C:\JOBST\DRAW.D"
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = FN_WHOpen(sFile, txtPatientName, txtFileNo, sSide)
		
		BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WH_LEG1.D;")
		
		'Put standard patient details
		PR_StandardDetails()
		
		'Body details
		BDYUTILS.PR_PutStringAssign("sTOSCir", txtTOSCir)
		'Fold Height (Height to fold of buttocks), Used revised if given
		nFoldHt = ARMEDDIA1.fnDisplaytoInches(Val(txtFoldHt.Text))
		If ARMEDDIA1.fnGetNumber(txtBody.Text, 4) > 0 And ARMEDDIA1.fnGetNumber(txtBody.Text, 7) > 0 Then
			nFoldHt = ARMEDDIA1.fnDisplaytoInches(ARMEDDIA1.fnGetNumber(txtBody.Text, 7))
		End If
		BDYUTILS.PR_PutNumberAssign("nFoldHt", nFoldHt)
		BDYUTILS.PR_PutStringAssign("sCrotchStyle", txtCrotchStyle)
		BDYUTILS.PR_PutStringAssign("sFabric", txtFabric)
		
		BDYUTILS.PR_PutNumberAssign("nLegStyle", nLegStyle)
		BDYUTILS.PR_PutNumberAssign("nLeftLegStyle", nLeftLegStyle)
		BDYUTILS.PR_PutNumberAssign("nRightLegStyle", nRightLegStyle)
		
		'Leg Details
		If sSide = "Left" Then
			BDYUTILS.PR_PutLine("LeftLeg = %true;")
			BDYUTILS.PR_PutLine("RightLeg = %false;")
			sLengths = txtLeftLengths.Text 'Use later to establish 1st & Last tapes
			'Heel value
			'Overide the LegStyle if there is no heel value
			nValue = Val(Mid(sLengths, (6 * 4) + 1, 4)) / 10
			
			If nLeftLegStyle <> 0 Or nValue = 0 Then
				BDYUTILS.PR_PutLine("FootLess = %true;")
				Footless = True
			Else
				BDYUTILS.PR_PutLine("FootLess = %false;")
				'Ankle tape values
				BDYUTILS.PR_PutNumberAssign("nAnkleTape", ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 1))
				BDYUTILS.PR_PutStringAssign("sAnkleTape", Str(ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 1)))
				BDYUTILS.PR_PutStringAssign("sMMAnkle", Str(ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 2)))
				BDYUTILS.PR_PutStringAssign("sGramsAnkle", Str(ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 3)))
				BDYUTILS.PR_PutStringAssign("sReductionAnkle", Str(ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 4)))
				BDYUTILS.PR_PutNumberAssign("nReductionAnkle", ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 4))
			End If
			
			nFabricClass = ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 8)
			If nFabricClass = 2 Then
				BDYUTILS.PR_PutStringAssign("sReduction", txtLeftRed)
				BDYUTILS.PR_PutStringAssign("sTapeMMs", txtLeftMMs)
				BDYUTILS.PR_PutStringAssign("sStretch", txtLeftStr)
			End If
			
			BDYUTILS.PR_PutStringAssign("sFirstLeg", "Left")
			BDYUTILS.PR_PutStringAssign("sLeg", "Left")
			BDYUTILS.PR_PutStringAssign("sSecondLeg", "Right")
			
			BDYUTILS.PR_PutStringAssign("sPressure", txtLeftTemplate)
			itemplate = Val(VB.Left(txtLeftTemplate.Text, 2))
			
			BDYUTILS.PR_PutStringAssign("sTapeLengths", txtLeftLengths)
			BDYUTILS.PR_PutStringAssign("sToeStyle", txtLeftToeStyle)
			
			BDYUTILS.PR_PutNumberAssign("nFootPleat1", ARMEDDIA1.fnDisplaytoInches(Val(txtLeftFootPleat1.Text)))
			BDYUTILS.PR_PutNumberAssign("nFootPleat2", ARMEDDIA1.fnDisplaytoInches(Val(txtLeftFootPleat2.Text)))
			BDYUTILS.PR_PutNumberAssign("nTopLegPleat1", ARMEDDIA1.fnDisplaytoInches(Val(txtLeftTopLegPleat1.Text)))
			BDYUTILS.PR_PutNumberAssign("nTopLegPleat2", ARMEDDIA1.fnDisplaytoInches(Val(txtLeftTopLegPleat2.Text)))
			
			BDYUTILS.PR_PutStringAssign("sFootLength", txtLeftFootLength)
			BDYUTILS.PR_PutNumberAssign("nElastic", ARMEDDIA1.fnGetNumber(txtLeftData.Text, 1))
			
			'Other leg details used to control toe
			nValue = ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 1)
			If nRightLegStyle <> 0 Or nValue < 0 Then
				BDYUTILS.PR_PutNumberAssign("nOtherAnkleMTapeLen ", 0)
			Else
				'Establish - Minus Ankle tape value
				If nValue = 8 Then
					'Ankle at +3
					nValue = Val(Mid(txtRightLengths.Text, (3 * 4) + 1, 4)) / 10
				Else
					'Ankle at +1-1/2
					nValue = Val(Mid(txtRightLengths.Text, (4 * 4) + 1, 4)) / 10
				End If
				BDYUTILS.PR_PutNumberAssign("nOtherAnkleMTapeLen ", ARMEDDIA1.fnDisplaytoInches(nValue))
			End If
			
		Else
			BDYUTILS.PR_PutLine("LeftLeg = %false;")
			BDYUTILS.PR_PutLine("RightLeg = %true;")
			sLengths = txtRightLengths.Text 'Use later to establish 1st & Last tapes
			'Heel value
			'Overide the LegStyle if there is no heel value
			nValue = Val(Mid(sLengths, (6 * 4) + 1, 4)) / 10
			
			If nRightLegStyle <> 0 Or nValue = 0 Then
				BDYUTILS.PR_PutLine("FootLess = %true;")
				Footless = True
			Else
				BDYUTILS.PR_PutLine("FootLess = %false;")
				'Ankle tape values
				BDYUTILS.PR_PutNumberAssign("nAnkleTape", ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 1))
				BDYUTILS.PR_PutStringAssign("sAnkleTape", Str(ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 1)))
				BDYUTILS.PR_PutStringAssign("sMMAnkle", Str(ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 2)))
				BDYUTILS.PR_PutStringAssign("sGramsAnkle", Str(ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 3)))
				BDYUTILS.PR_PutStringAssign("sReductionAnkle", Str(ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 4)))
				BDYUTILS.PR_PutNumberAssign("nReductionAnkle", ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 4))
			End If
			
			nFabricClass = ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 8)
			If nFabricClass = 2 Then
				BDYUTILS.PR_PutStringAssign("sReduction", txtRightRed)
				BDYUTILS.PR_PutStringAssign("sTapeMMs", txtRightMMs)
				BDYUTILS.PR_PutStringAssign("sStretch", txtRightStr)
			End If
			
			BDYUTILS.PR_PutStringAssign("sFirstLeg", "Right")
			BDYUTILS.PR_PutStringAssign("sLeg", "Right")
			BDYUTILS.PR_PutStringAssign("sSecondLeg", "Left")
			
			BDYUTILS.PR_PutStringAssign("sPressure", txtRightTemplate)
			itemplate = Val(VB.Left(txtRightTemplate.Text, 2))
			
			
			BDYUTILS.PR_PutStringAssign("sTapeLengths", txtRightLengths)
			BDYUTILS.PR_PutStringAssign("sToeStyle", txtRightToeStyle)
			
			BDYUTILS.PR_PutNumberAssign("nFootPleat1", ARMEDDIA1.fnDisplaytoInches(Val(txtRightFootPleat1.Text)))
			BDYUTILS.PR_PutNumberAssign("nFootPleat2", ARMEDDIA1.fnDisplaytoInches(Val(txtRightFootPleat2.Text)))
			BDYUTILS.PR_PutNumberAssign("nTopLegPleat1", ARMEDDIA1.fnDisplaytoInches(Val(txtRightTopLegPleat1.Text)))
			BDYUTILS.PR_PutNumberAssign("nTopLegPleat2", ARMEDDIA1.fnDisplaytoInches(Val(txtRightTopLegPleat2.Text)))
			
			BDYUTILS.PR_PutStringAssign("sFootLength", txtRightFootLength)
			BDYUTILS.PR_PutNumberAssign("nElastic", ARMEDDIA1.fnGetNumber(txtRightData.Text, 1))
			
			'Other leg details used to control toe
			nValue = ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 1)
			If nLeftLegStyle <> 0 Or nValue < 0 Then
				BDYUTILS.PR_PutNumberAssign("nOtherAnkleMTapeLen ", 0)
			Else
				'Establish - Minus Ankle tape value
				If nValue = 8 Then
					'Ankle at +3
					nValue = Val(Mid(txtLeftLengths.Text, (3 * 4) + 1, 4)) / 10
				Else
					'Ankle at +1-1/2
					nValue = Val(Mid(txtLeftLengths.Text, (4 * 4) + 1, 4)) / 10
				End If
				BDYUTILS.PR_PutNumberAssign("nOtherAnkleMTapeLen ", ARMEDDIA1.fnDisplaytoInches(nValue))
			End If
			
			
		End If
		
		'First and LastTapes
		nFirstTape = -1
		nLastTape = 30
		For ii = 0 To 29
			nValue = Val(Mid(sLengths, (ii * 4) + 1, 4)) / 10
			'Set first and last tape (assumes no holes in data)
			If nFirstTape < 0 And nValue > 0 Then nFirstTape = ii + 1
			If nLastTape = 30 And nFirstTape > 0 And nValue = 0 Then nLastTape = ii
		Next ii
		BDYUTILS.PR_PutNumberAssign("nFirstTape", nFirstTape)
		BDYUTILS.PR_PutNumberAssign("nLastTape", nLastTape)
		
		'Fabric Class (Load JOBSTEX_FL Procedures and Defaults if class = 2 )
		'For a footless style the only valid classes are 0 and 1
		If Footless And nFabricClass = 2 Then nFabricClass = 1
		BDYUTILS.PR_PutNumberAssign("nFabricClass", nFabricClass)
		
		
		'Draw leg except for BREIF
		If nLegStyle = 2 Then
			BDYUTILS.PR_PutStringAssign("sPressure", "-")
			'Draw Back body
			BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHLG1BRF.D;")
			
		Else
			'Load universal Procedures and Defaults
			BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHLG1DEF.D;")
			
			'Template type
			Select Case itemplate
				Case 13
					BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WH13DS.D;")
				Case 9
					BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WH09DS.D;")
				Case Else
					BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHPOW.D;")
			End Select
			
			
			'Get Origin
			BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHLG1ORG.D;")
			
			'Draw template
			BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHLGxTMP.D;")
			
			'Calculate Foot Points (If a foot exists)
			If Not Footless Then BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHFTPNTS.D;")
			
			'Draw Leg
			BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHLG1DWG.D;")
		End If
		
		'Close macro file
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
	End Sub
	
	Private Sub PR_WH_2ndLeg()
		'Procedure to create a Macro to draw the 2nd leg of a
		'Waist Height
		Dim sSide, sFile, sBody As String
		Dim sLengths As String
		Dim nLeftLegStyle, nLegStyle, nRightLegStyle As Short
		Dim nLastTape, nFirstTape, nFabricClass As Short
		Dim Footless, ii, itemplate As Short
		Dim nFoldHt, nValue As Double
		Dim nThighCir As Double
		
		'Establish which is the First Leg
		nLegStyle = ARMEDDIA1.fnGetNumber(txtLegStyle.Text, 1)
		Select Case nLegStyle
			Case 0, 1
				If ARMEDDIA1.fnDisplaytoInches(Val(txtLeftThighCir.Text)) - ARMEDDIA1.fnDisplaytoInches(Val(txtRightThighCir.Text)) > 1 Then
					sSide = "Left"
					nThighCir = ARMEDDIA1.fnDisplaytoInches(Val(txtRightThighCir.Text)) * Val(txtThighRed.Text) / 2
				Else
					sSide = "Right"
					nThighCir = ARMEDDIA1.fnDisplaytoInches(Val(txtLeftThighCir.Text)) * Val(txtThighRed.Text) / 2
				End If
			Case 2
				MsgBox("Brief Style Not yet supported", 16, "g_sDialogueID")
				End
			Case 3 To 6
				MsgBox("Can't draw a second leg for this style", 16, g_sDialogueID)
				End
		End Select
		
		
		'Open Macro file (fNum is declared as Global)
		sFile = "C:\JOBST\DRAW.D"
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = FN_WHOpen(sFile, txtPatientName, txtFileNo, sSide)
		
		BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WH_LEG2.D;")
		
		'Put standard patient details
		PR_StandardDetails()
		
		'Body details
		BDYUTILS.PR_PutStringAssign("sTOSCir", txtTOSCir)
		BDYUTILS.PR_PutNumberAssign("nTOSCir", Val(txtTOSCir.Text))
		
		'Fold Height (Height to fold of buttocks), Used revised if given
		nFoldHt = ARMEDDIA1.fnDisplaytoInches(Val(txtFoldHt.Text))
		sBody = txtBody.Text
		If ARMEDDIA1.fnGetNumber(sBody, 4) > 0 And ARMEDDIA1.fnGetNumber(sBody, 7) > 0 Then
			nFoldHt = ARMEDDIA1.fnDisplaytoInches(ARMEDDIA1.fnGetNumber(sBody, 7))
		End If
		
		BDYUTILS.PR_PutNumberAssign("nFoldHt", nFoldHt)
		BDYUTILS.PR_PutStringAssign("sCrotchStyle", txtCrotchStyle)
		BDYUTILS.PR_PutStringAssign("sFabric", txtFabric)
		
		nLegStyle = ARMEDDIA1.fnGetNumber(txtLegStyle.Text, 1)
		nLeftLegStyle = ARMEDDIA1.fnGetNumber(txtLegStyle.Text, 2)
		nRightLegStyle = ARMEDDIA1.fnGetNumber(txtLegStyle.Text, 3)
		
		BDYUTILS.PR_PutNumberAssign("nLegStyle", nLegStyle)
		BDYUTILS.PR_PutNumberAssign("nLeftLegStyle", nLeftLegStyle)
		BDYUTILS.PR_PutNumberAssign("nRightLegStyle", nRightLegStyle)
		
		
		'Leg Details
		If sSide = "Left" Then
			'The Cut-Out is always drawn for the Left side
			BDYUTILS.PR_PutLine("LeftLeg = %true;")
			BDYUTILS.PR_PutLine("RightLeg = %false;")
			sLengths = txtLeftLengths.Text 'Use later to establish 1st & Last tapes
			'Overide the LegStyle if there is no heel value
			nValue = Val(Mid(sLengths, (6 * 4) + 1, 4)) / 10
			If nLeftLegStyle <> 0 Or nValue = 0 Then
				BDYUTILS.PR_PutLine("FootLess = %true;")
				Footless = True
			Else
				BDYUTILS.PR_PutLine("FootLess = %false;")
				'Ankle tape values
				BDYUTILS.PR_PutNumberAssign("nAnkleTape", ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 1))
				BDYUTILS.PR_PutStringAssign("sAnkleTape", Str(ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 1)))
				BDYUTILS.PR_PutStringAssign("sMMAnkle", Str(ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 2)))
				BDYUTILS.PR_PutStringAssign("sGramsAnkle", Str(ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 3)))
				BDYUTILS.PR_PutStringAssign("sReductionAnkle", Str(ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 4)))
				BDYUTILS.PR_PutNumberAssign("nReductionAnkle", ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 4))
			End If
			
			nFabricClass = ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 8)
			If nFabricClass = 2 Then
				BDYUTILS.PR_PutStringAssign("sReduction", txtLeftRed)
				BDYUTILS.PR_PutStringAssign("sTapeMMs", txtLeftMMs)
				BDYUTILS.PR_PutStringAssign("sStretch", txtLeftStr)
			End If
			
			BDYUTILS.PR_PutStringAssign("sFirstLeg", "Right")
			BDYUTILS.PR_PutStringAssign("sLeg", "Left")
			BDYUTILS.PR_PutStringAssign("sSecondLeg", "Left")
			
			BDYUTILS.PR_PutStringAssign("sPressure", txtLeftTemplate)
			itemplate = Val(VB.Left(txtLeftTemplate.Text, 2))
			
			BDYUTILS.PR_PutStringAssign("sTapeLengths", txtLeftLengths)
			BDYUTILS.PR_PutStringAssign("sToeStyle", txtLeftToeStyle)
			
			BDYUTILS.PR_PutNumberAssign("nFootPleat1", ARMEDDIA1.fnDisplaytoInches(Val(txtLeftFootPleat1.Text)))
			BDYUTILS.PR_PutNumberAssign("nFootPleat2", ARMEDDIA1.fnDisplaytoInches(Val(txtLeftFootPleat2.Text)))
			BDYUTILS.PR_PutNumberAssign("nTopLegPleat1", ARMEDDIA1.fnDisplaytoInches(Val(txtLeftTopLegPleat1.Text)))
			BDYUTILS.PR_PutNumberAssign("nTopLegPleat2", ARMEDDIA1.fnDisplaytoInches(Val(txtLeftTopLegPleat2.Text)))
			
			BDYUTILS.PR_PutStringAssign("sFootLength", txtLeftFootLength)
			BDYUTILS.PR_PutNumberAssign("nElastic", ARMEDDIA1.fnGetNumber(txtLeftData.Text, 1))
			
			'Crotch style and data
			If txtCrotchStyle.Text = "Open Crotch" Then
				BDYUTILS.PR_PutLine("OpenCrotch  = %true;")
				BDYUTILS.PR_PutLine("ClosedCrotch = %false;")
			Else
				BDYUTILS.PR_PutLine("OpenCrotch  = %false;")
				BDYUTILS.PR_PutLine("ClosedCrotch = %true;")
			End If
			
			BDYUTILS.PR_PutNumberAssign("nCrotchFrontFactor", ARMEDDIA1.fnGetNumber(sBody, 1))
			BDYUTILS.PR_PutNumberAssign("nOpenFront", ARMEDDIA1.fnGetNumber(sBody, 2))
			BDYUTILS.PR_PutNumberAssign("nOpenBack", ARMEDDIA1.fnGetNumber(sBody, 3))
			
			'Other leg details used to control toe
			nValue = ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 1)
			If nRightLegStyle <> 0 Or nValue < 0 Then
				BDYUTILS.PR_PutNumberAssign("nOtherAnkleMTapeLen ", 0)
			Else
				'Establish - Minus Ankle tape value
				If nValue = 8 Then
					'Ankle at +3
					nValue = Val(Mid(txtRightLengths.Text, (3 * 4) + 1, 4)) / 10
				Else
					'Ankle at +1-1/2
					nValue = Val(Mid(txtRightLengths.Text, (4 * 4) + 1, 4)) / 10
				End If
				BDYUTILS.PR_PutNumberAssign("nOtherAnkleMTapeLen ", ARMEDDIA1.fnDisplaytoInches(nValue))
			End If
			
		Else
			BDYUTILS.PR_PutLine("LeftLeg = %false;")
			BDYUTILS.PR_PutLine("RightLeg = %true;")
			sLengths = txtRightLengths.Text 'Use later to establish 1st & Last tapes
			nValue = Val(Mid(sLengths, (6 * 4) + 1, 4)) / 10
			If nRightLegStyle <> 0 Or nValue = 0 Then
				BDYUTILS.PR_PutLine("FootLess = %true;")
				Footless = True
			Else
				BDYUTILS.PR_PutLine("FootLess = %false;")
				'Ankle tape values
				BDYUTILS.PR_PutNumberAssign("nAnkleTape", ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 1))
				BDYUTILS.PR_PutStringAssign("sAnkleTape", Str(ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 1)))
				BDYUTILS.PR_PutStringAssign("sMMAnkle", Str(ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 2)))
				BDYUTILS.PR_PutStringAssign("sGramsAnkle", Str(ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 3)))
				BDYUTILS.PR_PutStringAssign("sReductionAnkle", Str(ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 4)))
				BDYUTILS.PR_PutNumberAssign("nReductionAnkle", ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 4))
			End If
			
			nFabricClass = ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 8)
			If nFabricClass = 2 Then
				BDYUTILS.PR_PutStringAssign("sReduction", txtRightRed)
				BDYUTILS.PR_PutStringAssign("sTapeMMs", txtRightMMs)
				BDYUTILS.PR_PutStringAssign("sStretch", txtRightStr)
			End If
			
			BDYUTILS.PR_PutStringAssign("sFirstLeg", "Left")
			BDYUTILS.PR_PutStringAssign("sLeg", "Right")
			BDYUTILS.PR_PutStringAssign("sSecondLeg", "Right")
			
			BDYUTILS.PR_PutStringAssign("sPressure", txtRightTemplate)
			itemplate = Val(VB.Left(txtRightTemplate.Text, 2))
			
			BDYUTILS.PR_PutStringAssign("sTapeLengths", txtRightLengths)
			BDYUTILS.PR_PutStringAssign("sToeStyle", txtRightToeStyle)
			
			BDYUTILS.PR_PutNumberAssign("nFootPleat1", ARMEDDIA1.fnDisplaytoInches(Val(txtRightFootPleat1.Text)))
			BDYUTILS.PR_PutNumberAssign("nFootPleat2", ARMEDDIA1.fnDisplaytoInches(Val(txtRightFootPleat2.Text)))
			BDYUTILS.PR_PutNumberAssign("nTopLegPleat1", ARMEDDIA1.fnDisplaytoInches(Val(txtRightTopLegPleat1.Text)))
			BDYUTILS.PR_PutNumberAssign("nTopLegPleat2", ARMEDDIA1.fnDisplaytoInches(Val(txtRightTopLegPleat2.Text)))
			
			BDYUTILS.PR_PutStringAssign("sFootLength", txtRightFootLength)
			BDYUTILS.PR_PutNumberAssign("nElastic", ARMEDDIA1.fnGetNumber(txtRightData.Text, 1))
			
			'Other leg details used to control toe
			nValue = ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 1)
			If nLeftLegStyle <> 0 Or nValue < 0 Then
				BDYUTILS.PR_PutNumberAssign("nOtherAnkleMTapeLen ", 0)
			Else
				'Establish - Minus Ankle tape value
				If nValue = 8 Then
					'Ankle at +3
					nValue = Val(Mid(txtLeftLengths.Text, (3 * 4) + 1, 4)) / 10
				Else
					'Ankle at +1-1/2
					nValue = Val(Mid(txtLeftLengths.Text, (4 * 4) + 1, 4)) / 10
				End If
				BDYUTILS.PR_PutNumberAssign("nOtherAnkleMTapeLen ", ARMEDDIA1.fnDisplaytoInches(nValue))
			End If
			
			
		End If
		
		'First and LastTapes
		nFirstTape = -1
		nLastTape = 30
		For ii = 0 To 29
			nValue = Val(Mid(sLengths, (ii * 4) + 1, 4)) / 10
			'Set first and last tape (assumes no holes in data)
			If nFirstTape < 0 And nValue > 0 Then nFirstTape = ii + 1
			If nLastTape = 30 And nFirstTape > 0 And nValue = 0 Then nLastTape = ii
		Next ii
		BDYUTILS.PR_PutNumberAssign("nFirstTape", nFirstTape)
		BDYUTILS.PR_PutNumberAssign("nLastTape", nLastTape)
		
		'Actual Waist and Figured thigh Circumference
		'Used for body and lateral zips
		BDYUTILS.PR_PutNumberAssign("nGivenWaistCir", ARMEDDIA1.fnDisplaytoInches(Val(txtWaistCir.Text)))
		BDYUTILS.PR_PutNumberAssign("nThighCir", nThighCir + 0.1875)
		
		'Fabric Class (Load JOBSTEX_FL Procedures and Defaults if class = 2 )
		'For a footless style the only valid classes are 0 and 1
		If Footless And nFabricClass = 2 Then nFabricClass = 1
		BDYUTILS.PR_PutNumberAssign("nFabricClass", nFabricClass)
		
		'Load universal Procedures and Defaults
		BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHLG2DEF.D;")
		
		'Get 1st leg profile
		BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHLG2LEG.D;")
		
		'Template type
		Select Case itemplate
			Case 13
				BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WH13DS.D;")
			Case 9
				BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WH09DS.D;")
			Case Else
				BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHPOW.D;")
		End Select
		
		'Get Origin
		BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHLG2ORG.D;")
		
		'Draw template
		BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHLGxTMP.D;")
		
		'Draw Body
		If sSide = "Left" Then
			BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHLG2CUT.D;")
		Else
			BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHLG2DBD.D;")
		End If
		
		'Calculate Foot Points (If a foot exists)
		If Not Footless Then BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHFTPNTS.D;")
		
		'Draw Leg
		BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHLG2DWG.D;")
		
		'Close macro file
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
	End Sub
	
	Private Sub PR_WH_Brief1stLeg()
		'Procedure to create a Macro to draw the first leg of a
		'Waist Height
		Dim sFile, sSide As String
		Dim sLengths As String
		Dim nLeftLegStyle, nLegStyle, nRightLegStyle As Short
		Dim nLastTape, nFirstTape, nFabricClass As Short
		Dim Footless, ii, itemplate As Short
		Dim nFoldHt, nValue As Double
		
		nLegStyle = ARMEDDIA1.fnGetNumber(txtLegStyle.Text, 1)
		sFile = "C:\JOBST\DRAW.D"
		
		'For a breif the side is always LEFT
		sSide = "Left"
		
		'Open Macro file (fNum is declared as Global)
		sFile = "C:\JOBST\DRAW.D"
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = FN_WHOpen(sFile, txtPatientName, txtFileNo, sSide)
		
		BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WH_LEG1.D;")
		
		'Put standard patient details
		PR_StandardDetails()
		
		'Body details
		BDYUTILS.PR_PutStringAssign("sTOSCir", txtTOSCir)
		
		'Fold Height (Height to fold of buttocks), Used revised if given
		nFoldHt = ARMEDDIA1.fnDisplaytoInches(Val(txtFoldHt.Text))
		If ARMEDDIA1.fnGetNumber(txtBody.Text, 4) > 0 And ARMEDDIA1.fnGetNumber(txtBody.Text, 7) > 0 Then
			nFoldHt = ARMEDDIA1.fnDisplaytoInches(ARMEDDIA1.fnGetNumber(txtBody.Text, 7))
		End If
		
		BDYUTILS.PR_PutNumberAssign("nFoldHt", nFoldHt)
		BDYUTILS.PR_PutStringAssign("sCrotchStyle", txtCrotchStyle)
		BDYUTILS.PR_PutStringAssign("sFabric", txtFabric)
		
		nLegStyle = ARMEDDIA1.fnGetNumber(txtLegStyle.Text, 1)
		
		'Leg Details
		BDYUTILS.PR_PutNumberAssign("nLegStyle", nLegStyle)
		BDYUTILS.PR_PutLine("LeftLeg = %true;")
		BDYUTILS.PR_PutLine("RightLeg = %false;")
		BDYUTILS.PR_PutLine("FootLess = %true;")
		BDYUTILS.PR_PutStringAssign("sFirstLeg", "Left")
		BDYUTILS.PR_PutStringAssign("sLeg", "Left")
		BDYUTILS.PR_PutStringAssign("sSecondLeg", "Right")
		BDYUTILS.PR_PutNumberAssign("nFabricClass", 0)
		BDYUTILS.PR_PutStringAssign("sPressure", "-")
		
		'Draw Back body
		BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHLG1BRF.D;")
		
		'Close macro file
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
	End Sub
	
	Private Sub PR_WH_Brief2ndLeg()
		'Procedure to create a Macro to draw the Second leg of a
		'Waist Height Breif
		Dim sFile, sSide As String
		Dim sLengths As String
		Dim nLeftLegStyle, nLegStyle, nRightLegStyle As Short
		Dim nLastTape, nFirstTape, nFabricClass As Short
		Dim Footless, ii, itemplate As Short
		Dim nFoldHt, nValue As Double
		
		nLegStyle = ARMEDDIA1.fnGetNumber(txtLegStyle.Text, 1)
		sFile = "C:\JOBST\DRAW.D"
		
		'For a breif the side is always LEFT
		sSide = "Left"
		
		'Open Macro file (fNum is declared as Global)
		sFile = "C:\JOBST\DRAW.D"
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = FN_WHOpen(sFile, txtPatientName, txtFileNo, sSide)
		
		BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WH_LEG2.D;")
		
		'Put standard patient details
		PR_StandardDetails()
		
		'Body details
		BDYUTILS.PR_PutStringAssign("sTOSCir", txtTOSCir)
		
		'Fold Height (Height to fold of buttocks), Used revised if given
		nFoldHt = ARMEDDIA1.fnDisplaytoInches(Val(txtFoldHt.Text))
		If ARMEDDIA1.fnGetNumber(txtBody.Text, 4) > 0 And ARMEDDIA1.fnGetNumber(txtBody.Text, 7) > 0 Then
			nFoldHt = ARMEDDIA1.fnDisplaytoInches(ARMEDDIA1.fnGetNumber(txtBody.Text, 7))
		End If
		
		BDYUTILS.PR_PutNumberAssign("nFoldHt", nFoldHt)
		BDYUTILS.PR_PutStringAssign("sCrotchStyle", txtCrotchStyle)
		BDYUTILS.PR_PutStringAssign("sFabric", txtFabric)
		
		nLegStyle = ARMEDDIA1.fnGetNumber(txtLegStyle.Text, 1)
		
		'Leg Details
		BDYUTILS.PR_PutNumberAssign("nLegStyle", nLegStyle)
		BDYUTILS.PR_PutLine("LeftLeg = %true;")
		BDYUTILS.PR_PutLine("RightLeg = %false;")
		BDYUTILS.PR_PutLine("FootLess = %true;")
		BDYUTILS.PR_PutStringAssign("sLeg", "Left")
		BDYUTILS.PR_PutStringAssign("sFirstLeg", "Left")
		BDYUTILS.PR_PutStringAssign("sSecondLeg", "Right")
		BDYUTILS.PR_PutNumberAssign("nFabricClass", 0)
		BDYUTILS.PR_PutStringAssign("sPressure", "-")
		
		If txtCrotchStyle.Text = "Open Crotch" Then
			BDYUTILS.PR_PutLine("OpenCrotch = %true;")
			BDYUTILS.PR_PutNumberAssign("nBriefOff", 1.625)
		Else
			BDYUTILS.PR_PutLine("OpenCrotch = %false;")
			BDYUTILS.PR_PutNumberAssign("nBriefOff", 1.25)
		End If
		
		'Fold offset
		If txtCrotchStyle.Text = "Horizontal Fly" Then
			BDYUTILS.PR_PutNumberAssign("nFoldOff", 0)
		Else
			BDYUTILS.PR_PutNumberAssign("nFoldOff", 0.75)
		End If
		
		'Draw Back body
		BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHLG2LEG.D;")
		BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHLG2BRF.D;")
		
		'Close macro file
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
		
	End Sub
	
	Private Sub PR_WH_ChapBody(ByRef sSide As String)
		'Procedure to draw the chap style BODY
		Dim xyStartBotOFF, xyStartCL, xyPt2, xyPt, xyPt1, xyPt3, xyStartOFF, xyInt As xy
		Dim xyWaistBandProx, xyEOSCL, xyChapDatumOFF, xyWaistBandDist As xy
		Dim xyVelcroOverLapProx, xyVelcroOverLapDist As xy
		Dim iDistCase, iProxCase As Short
		Dim xyChapBotDatumOFF As xy
		Dim xyCenTop, xyCenTopProx, xyCenTopDist, xyArcBottomPt As xy
		Dim xyTmpDist, xyCenBotDist, xyCenBotProx, xyCenBot, xyTmpProx As xy
		Dim aAngle, nTopRaduisDist, nTopRadius, nTopRaduisProx, nLength, nEighthOfWaistCir As Double
		Dim nBotRaduisProx, nBotRadius, nBotRaduisDist As Double
		'UPGRADE_WARNING: Arrays in structure TmpBotlegProfile may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim TmpBotlegProfile As Curve
		Dim aTheta, nB, nA, nC, nCosTheta As Double
		Dim ii, iStart, Sucess As Short
		Dim nTOSHt, nWaistHt, nWaistCir, nTOSCir, nFoldHt, nReduction As Double
		Dim sLegFile As String
		Dim fFile As Short
		Dim nn As Double
		
		
		'Get Leg curve from file
		If FileLen("C:\JOBST\LEGCURVE.DAT") = 0 Then
			MsgBox("C:\JOBST\LEGCURVE.DAT" & "Not found", 48, g_sDialogueID)
			Exit Sub
		End If
		
		fFile = FreeFile
		FileOpen(fFile, "C:\JOBST\LEGCURVE.DAT", OpenMode.Input)
		
		'Get control points
		Input(fFile, xyOtemplate.X)
		Input(fFile, xyOtemplate.Y)
		
		'Get profile points
		g_TopLegProfile.n = 0
		Do While Not EOF(fFile)
			g_TopLegProfile.n = g_TopLegProfile.n + 1
			Input(fFile, g_TopLegProfile.X(g_TopLegProfile.n))
			Input(fFile, g_TopLegProfile.Y(g_TopLegProfile.n))
		Loop 
		FileClose(fFile)
		
		
		'Revise xyChapDatum wrt leg and last tape
		xyChapDatum.X = g_TopLegProfile.X(g_TopLegProfile.n) - 0.75
		xyChapDatum.Y = xyOtemplate.Y
		
		'Get the leg points that straddle x = xyChapDatum.x.
		'This is done to take account of pleats
		
		If g_TopLegProfile.n - 5 <= 0 Then iStart = 1 Else iStart = g_TopLegProfile.n - 5
		
		'From top
		For nn = g_TopLegProfile.n To iStart Step -1
			If g_TopLegProfile.X(nn) > xyChapDatum.X Then
				xyProfileProximal.X = g_TopLegProfile.X(nn)
				xyProfileProximal.Y = g_TopLegProfile.Y(nn)
			End If
		Next nn
		
		'From bottom
		'NB
		' 1. We also create the mirrored bottom leg profile
		' 2. Special case where g_ToplegProfile.X(nn) = xyChapDatum.X
		g_BotLegProfile.n = 0
		For nn = iStart To g_TopLegProfile.n
			g_BotLegProfile.n = g_BotLegProfile.n + 1
			g_BotLegProfile.X(g_BotLegProfile.n) = g_TopLegProfile.X(nn)
			nLength = g_TopLegProfile.Y(nn) - xyChapDatum.Y
			g_BotLegProfile.Y(g_BotLegProfile.n) = xyChapDatum.Y - nLength
			If g_TopLegProfile.X(nn) <= xyChapDatum.X Then
				If g_TopLegProfile.X(nn) = xyChapDatum.X Then
					xyProfileDistal.X = g_TopLegProfile.X(nn) - 0.001
				Else
					xyProfileDistal.X = g_TopLegProfile.X(nn)
				End If
				xyProfileDistal.Y = g_TopLegProfile.Y(nn)
			End If
		Next nn
		
		'Circumferences
		nTOSCir = ARMEDDIA1.fnDisplaytoInches(Val(txtTOSCir.Text))
		nWaistCir = ARMEDDIA1.fnDisplaytoInches(Val(txtWaistCir.Text))
		nTOSHt = ARMEDDIA1.fnDisplaytoInches(Val(txtTOSHt.Text))
		nFoldHt = ARMEDDIA1.fnDisplaytoInches(Val(txtFoldHt.Text))
		nWaistHt = ARMEDDIA1.fnDisplaytoInches(Val(txtWaistHt.Text))
		If nTOSCir <> 0 And nTOSCir < nWaistCir Then
			nReduction = Val(txtTOSRed.Text)
			g_nChapCirGiven = nTOSCir
			g_nChapCirFigured = (nTOSCir * nReduction) / 2
			g_nChapLenGiven = nTOSHt - nFoldHt
		Else
			nReduction = Val(txtWaistRed.Text)
			g_nChapCirGiven = nWaistCir
			g_nChapCirFigured = (nWaistCir * nReduction) / 2
			g_nChapLenGiven = nWaistHt - nFoldHt
		End If
		g_nChapLenFigured = (g_nChapLenGiven / 1.2) + 0.75
		
		g_nLegStyle = ARMEDDIA1.fnGetNumber(txtLegStyle.Text, 1)
		
		
		'Get intersections on leg profile
		'Note special case for only 3 points in leg profile
		If g_TopLegProfile.n - 5 <= 0 Then iStart = 1 Else iStart = g_TopLegProfile.n - 5
		
		'PR_MakeXY xyOrigin, 0, xyChapDatum.y
		'UPGRADE_WARNING: Couldn't resolve default property of object xyOrigin. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xyOrigin = xyOtemplate
		ARMDIA1.PR_MakeXY(xyStartCL, xyChapDatum.X - 2.5, xyChapDatum.Y)
		ARMDIA1.PR_MakeXY(xyStartOFF, xyStartCL.X, xyChapDatum.Y + 100)
		
		ARMDIA1.PR_MakeXY(xyPt1, g_TopLegProfile.X(iStart), g_TopLegProfile.Y(iStart))
		
		Sucess = False
		If g_TopLegProfile.n = 3 Then iStart = 0 'NB 2 tape leg
		For ii = iStart + 1 To g_TopLegProfile.n
			ARMDIA1.PR_MakeXY(xyPt2, g_TopLegProfile.X(ii), g_TopLegProfile.Y(ii))
			If BDYUTILS.FN_LinLinInt(xyStartCL, xyStartOFF, xyPt1, xyPt2, xyStartOFF) Then
				Sucess = True
				Exit For
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object xyPt1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyPt1 = xyPt2
		Next ii
		
		'assume a short 3 tape leg
		'and extend the bottom leg profile
		If Not Sucess Then
			ARMDIA1.PR_MakeXY(xyStartOFF, xyStartCL.X, g_TopLegProfile.Y(1))
			
			For ii = 1 To g_BotLegProfile.n
				TmpBotlegProfile.X(ii + 1) = g_BotLegProfile.X(ii)
				TmpBotlegProfile.Y(ii + 1) = g_BotLegProfile.Y(ii)
			Next ii
			
			nLength = xyStartOFF.Y - xyChapDatum.Y
			TmpBotlegProfile.Y(1) = xyChapDatum.Y - nLength
			TmpBotlegProfile.X(1) = xyStartOFF.X
			TmpBotlegProfile.n = g_BotLegProfile.n + 1
			
			'UPGRADE_WARNING: Couldn't resolve default property of object g_BotLegProfile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_BotLegProfile = TmpBotlegProfile
			
		End If
		
		'Get intesection at X = xyChapDatum.X
		Sucess = False
		ARMDIA1.PR_MakeXY(xyChapDatumOFF, xyChapDatum.X, xyChapDatum.Y + 100)
		If BDYUTILS.FN_LinLinInt(xyChapDatum, xyChapDatumOFF, xyProfileDistal, xyProfileProximal, xyChapDatumOFF) Then
			Sucess = True
		End If
		
		'Give an error message here if we can't find a leg intesection
		If Not Sucess Then
			MsgBox("Can't form upper part of back edge of chap. Unable to find an intesection on the leg profile", 16, g_sDialogueID)
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileClose(fNum)
			End
		End If
		
		'Calculate key points
		nEighthOfWaistCir = g_nChapCirFigured / 4
		ARMDIA1.PR_MakeXY(xyEOSCL, xyChapDatum.X + g_nChapLenFigured, xyChapDatum.Y)
		ARMDIA1.PR_MakeXY(xyVelcroOverLapProx, xyEOSCL.X, xyChapDatum.Y + nEighthOfWaistCir)
		ARMDIA1.PR_MakeXY(xyVelcroOverLapDist, xyEOSCL.X - 1, xyChapDatum.Y + nEighthOfWaistCir)
		
		ARMDIA1.PR_MakeXY(xyWaistBandProx, xyVelcroOverLapProx.X, xyVelcroOverLapProx.Y - g_nChapCirFigured)
		ARMDIA1.PR_MakeXY(xyWaistBandDist, xyVelcroOverLapDist.X, xyVelcroOverLapDist.Y - g_nChapCirFigured)
		
		
		'Draw CHAP
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = FN_ChapOpen("C:\JOBST\DRAW.D")
		
		ARMDIA1.PR_SetLayer("Template" & sSide)
		ARMDIA1.PR_DrawLine(xyWaistBandDist, xyWaistBandProx)
		ARMDIA1.PR_DrawLine(xyWaistBandProx, xyVelcroOverLapProx)
		
		'Draw velcro OVERLAP
		'If two legs and we are drawing the Left then don't draw the
		'velcro overlap
		If g_nLegStyle = 7 And sSide = "Left" Then
			ARMDIA1.PR_DrawLine(xyVelcroOverLapProx, xyVelcroOverLapDist)
		Else
			'Label overlap so it can be deleted by the Zipper programme
			ARMDIA1.PR_MakeXY(xyPt1, xyVelcroOverLapProx.X, xyVelcroOverLapProx.Y + 2.5)
			ARMDIA1.PR_MakeXY(xyPt2, xyVelcroOverLapDist.X, xyVelcroOverLapDist.Y + 2.5)
			ARMDIA1.PR_DrawLine(xyVelcroOverLapProx, xyPt1)
			BDYUTILS.PR_AddDBValueToLast("Zipper", "VelcroOverlap")
			ARMDIA1.PR_DrawLine(xyPt1, xyPt2)
			BDYUTILS.PR_AddDBValueToLast("Zipper", "VelcroOverlap")
			ARMDIA1.PR_DrawLine(xyVelcroOverLapDist, xyPt2)
			BDYUTILS.PR_AddDBValueToLast("Zipper", "VelcroOverlap")
			
			ARMDIA1.PR_SetLayer("Notes")
			ARMDIA1.PR_DrawLine(xyVelcroOverLapProx, xyVelcroOverLapDist)
			BDYUTILS.PR_AddDBValueToLast("Zipper", "VelcroOverlap")
			BDYUTILS.PR_CalcMidPoint(xyPt1, xyVelcroOverLapDist, xyPt3)
			ARMDIA1.PR_InsertSymbol("TextAsSymbol", xyPt3, 1, 90)
			BDYUTILS.PR_AddDBValueToLast("Data", "2-1/2\"" VELCRO")
			BDYUTILS.PR_AddDBValueToLast("Zipper", "VelcroOverlap")
			ARMDIA1.PR_SetLayer("Template" & sSide)
			
		End If
		
		'Add One waist band for Chap Both Legs
		If g_nLegStyle = 7 Then
			ARMDIA1.PR_SetLayer("Notes")
			ARMDIA1.PR_SetTextData(HORIZ_CENTER, VERT_CENTER, 0.125, CURRENT, CURRENT)
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextAngle = 90
			BDYUTILS.PR_CalcMidPoint(xyEOSCL, xyVelcroOverLapProx, xyPt3)
			xyPt3.X = xyPt3.X - 0.5
			ARMDIA1.PR_DrawText("ONE WAIST BAND", xyPt3, 0.125)
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextAngle = 0
			ARMDIA1.PR_SetLayer("Template" & sSide)
		End If
		
		
		'Get the first aproximation for the Top and Bottom arcs
		'Note we work on the +ve side of the line Y = xyChapDatum.Y
		'and then mirror the results for the bottom arc
		
		'Top arc/s
		nTopRadius = (xyVelcroOverLapDist.X - xyChapDatum.X) / 2
		ARMDIA1.PR_MakeXY(xyCenTop, xyChapDatum.X + nTopRadius, xyChapDatum.Y + nTopRadius)
		
		'Adjust to fit if the center of the arc is above the velcro line
		'or the end of the profile
		ARMDIA1.PR_MakeXY(xyArcBottomPt, xyCenTop.X, xyChapDatum.Y)
		If xyCenTop.Y > xyVelcroOverLapDist.Y Then
			'Revise center of arc and Raduis
			ARMDIA1.PR_MakeXY(xyArcBottomPt, xyCenTop.X, xyChapDatum.Y)
			nLength = ARMDIA1.FN_CalcLength(xyArcBottomPt, xyVelcroOverLapDist) / 2
			aAngle = 90 - ARMDIA1.FN_CalcAngle(xyArcBottomPt, xyVelcroOverLapDist)
			nTopRaduisProx = nLength / System.Math.Cos(aAngle * (PI / 180))
			ARMDIA1.PR_MakeXY(xyCenTopProx, xyArcBottomPt.X, xyArcBottomPt.Y + nTopRaduisProx)
			ARMDIA1.PR_DrawArc(xyCenTopProx, xyArcBottomPt, xyVelcroOverLapDist)
		Else
			ARMDIA1.PR_MakeXY(xyPt1, xyCenTop.X + nTopRadius, xyCenTop.Y)
			ARMDIA1.PR_DrawLine(xyVelcroOverLapDist, xyPt1)
			ARMDIA1.PR_DrawArc(xyCenTop, xyArcBottomPt, xyPt1)
		End If
		
		If xyCenTop.Y > xyChapDatumOFF.Y Then
			'Revise center of arc and Raduis
			nLength = ARMDIA1.FN_CalcLength(xyArcBottomPt, xyChapDatumOFF) / 2
			aAngle = ARMDIA1.FN_CalcAngle(xyArcBottomPt, xyChapDatumOFF) - 90
			nTopRaduisDist = nLength / System.Math.Cos(aAngle * (PI / 180))
			ARMDIA1.PR_MakeXY(xyCenTopDist, xyArcBottomPt.X, xyArcBottomPt.Y + nTopRaduisDist)
			ARMDIA1.PR_DrawArc(xyCenTopDist, xyChapDatumOFF, xyArcBottomPt)
		Else
			ARMDIA1.PR_MakeXY(xyPt1, xyCenTop.X - nTopRadius, xyCenTop.Y)
			ARMDIA1.PR_DrawLine(xyChapDatumOFF, xyPt1)
			ARMDIA1.PR_DrawArc(xyCenTop, xyPt1, xyArcBottomPt)
		End If
		
		
		'Bottom arc/s
		'Calculate for +ve Y and then mirror results
		ARMDIA1.PR_MakeXY(xyArcBottomPt, xyCenTop.X, xyWaistBandDist.Y + nEighthOfWaistCir)
		
		
		'Mirror our two control points in the line Y = xyChapDatum.Y
		xyArcBottomPt.Y = xyChapDatum.Y + System.Math.Abs(xyArcBottomPt.Y - xyChapDatum.Y)
		xyWaistBandDist.Y = xyChapDatum.Y + System.Math.Abs(xyWaistBandDist.Y - xyChapDatum.Y)
		
		nBotRadius = nTopRadius
		ARMDIA1.PR_MakeXY(xyCenBot, xyArcBottomPt.X, xyArcBottomPt.Y + nBotRadius)
		
		If xyCenBot.Y > xyWaistBandDist.Y Then
			nLength = ARMDIA1.FN_CalcLength(xyArcBottomPt, xyWaistBandDist) / 2
			aAngle = ARMDIA1.FN_CalcAngle(xyArcBottomPt, xyWaistBandDist) - 90
			nBotRaduisProx = nLength / System.Math.Cos(aAngle * (PI / 180))
			ARMDIA1.PR_MakeXY(xyCenBotProx, xyArcBottomPt.X, xyArcBottomPt.Y + nBotRaduisProx)
			'UPGRADE_WARNING: Couldn't resolve default property of object xyTmpProx. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyTmpProx = xyWaistBandDist
			iProxCase = 1
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object xyCenBotProx. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyCenBotProx = xyCenBot
			nBotRaduisProx = nBotRadius
			ARMDIA1.PR_MakeXY(xyTmpProx, xyCenBot.X + nBotRadius, xyCenBot.Y)
			iProxCase = 2
		End If
		
		If xyCenBot.Y < xyChapDatumOFF.Y Then
			ARMDIA1.PR_MakeXY(xyTmpDist, xyCenBot.X - nBotRadius, xyCenBot.Y)
			'UPGRADE_WARNING: Couldn't resolve default property of object xyCenBotDist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xyCenBotDist = xyCenBot
			iDistCase = 1
			
		ElseIf xyArcBottomPt.Y < xyChapDatumOFF.Y Then 
			'Revise center of arc and Raduis
			nLength = ARMDIA1.FN_CalcLength(xyArcBottomPt, xyChapDatumOFF) / 2
			aAngle = 90 - ARMDIA1.FN_CalcAngle(xyArcBottomPt, xyChapDatumOFF)
			nBotRaduisDist = nLength / System.Math.Cos(aAngle * (PI / 180))
			ARMDIA1.PR_MakeXY(xyCenBotDist, xyArcBottomPt.X, xyArcBottomPt.Y + nBotRaduisDist)
			iDistCase = 2
		ElseIf xyArcBottomPt.Y = xyChapDatumOFF.Y Then 
			iDistCase = 3
		Else
			nA = ARMDIA1.FN_CalcLength(xyCenBotProx, xyChapDatumOFF)
			nC = nBotRaduisProx
			nB = System.Math.Sqrt(System.Math.Abs(nA ^ 2 - nC ^ 2))
			'Using the cosine rule to get the angle
			nCosTheta = (nC ^ 2 - (nA ^ 2 + nB ^ 2)) / -(2 * nA * nB)
			aTheta = BDYUTILS.Arccos(nCosTheta) * (180 / PI)
			aAngle = ARMDIA1.FN_CalcAngle(xyChapDatumOFF, xyCenBotProx) - aTheta
			ARMDIA1.PR_CalcPolar(xyChapDatumOFF, aAngle, nB, xyArcBottomPt)
			iDistCase = 4
		End If
		
		'Mirror our points in the line Y = xyChapDatum.Y
		xyArcBottomPt.Y = xyChapDatum.Y - System.Math.Abs(xyArcBottomPt.Y - xyChapDatum.Y)
		xyWaistBandDist.Y = xyChapDatum.Y - System.Math.Abs(xyWaistBandDist.Y - xyChapDatum.Y)
		xyCenBotProx.Y = xyChapDatum.Y - System.Math.Abs(xyCenBotProx.Y - xyChapDatum.Y)
		xyCenBotDist.Y = xyChapDatum.Y - System.Math.Abs(xyCenBotDist.Y - xyChapDatum.Y)
		'UPGRADE_WARNING: Couldn't resolve default property of object xyChapBotDatumOFF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xyChapBotDatumOFF = xyChapDatumOFF
		xyChapBotDatumOFF.Y = xyChapDatum.Y - System.Math.Abs(xyChapBotDatumOFF.Y - xyChapDatum.Y)
		'UPGRADE_WARNING: Couldn't resolve default property of object xyStartBotOFF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xyStartBotOFF = xyStartOFF
		xyStartBotOFF.Y = xyChapDatum.Y - System.Math.Abs(xyStartBotOFF.Y - xyChapDatum.Y)
		xyTmpProx.Y = xyChapDatum.Y - System.Math.Abs(xyTmpProx.Y - xyChapDatum.Y)
		xyTmpDist.Y = xyChapDatum.Y - System.Math.Abs(xyTmpDist.Y - xyChapDatum.Y)
		
		'Draw the arcs and lines based on the draw case
		If iProxCase = 2 Then ARMDIA1.PR_DrawLine(xyWaistBandDist, xyTmpProx)
		ARMDIA1.PR_DrawArc(xyCenBotProx, xyArcBottomPt, xyTmpProx)
		
		Select Case iDistCase
			Case 1
				ARMDIA1.PR_DrawLine(xyChapBotDatumOFF, xyTmpDist)
				ARMDIA1.PR_DrawArc(xyCenBotDist, xyTmpDist, xyArcBottomPt)
			Case 2
				ARMDIA1.PR_DrawArc(xyCenBotDist, xyChapBotDatumOFF, xyArcBottomPt)
			Case 3, 4
				ARMDIA1.PR_DrawLine(xyArcBottomPt, xyChapBotDatumOFF)
		End Select
		
		'Bottom leg profile
		ARMDIA1.PR_DrawFitted(g_BotLegProfile)
		ARMDIA1.PR_DrawLine(xyStartCL, xyEOSCL)
		
		'Draw markers for use with zippers
		ARMDIA1.PR_SetLayer("Construct")
		BDYUTILS.PR_DrawMarker(xyChapDatum, "Fold")
		
		BDYUTILS.PR_DrawMarker(xyEOSCL, "EOS")
		
		'Draw Closing lines
		ARMDIA1.PR_SetLayer("Template" & sSide)
		ARMDIA1.PR_DrawLine(xyStartCL, xyStartBotOFF)
		ARMDIA1.PR_DrawLine(xyOrigin, xyStartCL)
		
		'Tram lines
		ARMDIA1.PR_SetLayer("Notes")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object xyPt1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xyPt1 = xyOrigin
		xyPt1.Y = xyOrigin.Y + 0.1875
		'UPGRADE_WARNING: Couldn't resolve default property of object xyPt2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xyPt2 = xyStartCL
		xyPt2.Y = xyStartCL.Y + 0.1875
		ARMDIA1.PR_DrawLine(xyPt1, xyPt2)
		
		xyPt1.Y = xyPt2.Y + 0.5
		xyPt2.Y = xyPt2.Y + 0.5
		ARMDIA1.PR_DrawLine(xyPt1, xyPt2)
		
		
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
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
	End Sub
	
	Private Sub PR_WH_ChapLeg(ByRef sDrawType As String)
		'Draw a chap style
		Dim sFile As String
		Dim sLengths As String
		Dim nLeftLegStyle, nLegStyle, nRightLegStyle As Short
		Dim nFirstTape, nLastTape As Short
		Dim Footless, ii, itemplate As Short
		Dim nFoldHt, nValue As Double
		
		nLegStyle = ARMEDDIA1.fnGetNumber(txtLegStyle.Text, 1)
		g_nLegStyle = nLegStyle
		nLeftLegStyle = ARMEDDIA1.fnGetNumber(txtLegStyle.Text, 2)
		nRightLegStyle = ARMEDDIA1.fnGetNumber(txtLegStyle.Text, 3)
		
		sFile = "C:\JOBST\DRAW.D"
		
		'Establish which is the First Leg
		Select Case nLegStyle
			Case 5
				g_sSide = "Left"
			Case 6
				g_sSide = "Right"
			Case 7
				If sDrawType = "WH2LG" Then
					g_sSide = "Right"
				Else
					g_sSide = "Left"
				End If
		End Select
		
		'Having decided on leg do a swift check to see if there is enough
		'data to draw the leg
		If g_sSide = "Left" And ((txtLeftAnkle.Text = "" And txtLeftLengths.Text = "") Or txtUidLeftLeg.Text = "") Then
			MsgBox("No data for LEFT Leg", 16, g_sDialogueID)
			End
		End If
		
		If g_sSide = "Left" And txtLeftTemplate.Text = "" Then
			MsgBox("Figure LEFT Leg before drawing", 16, g_sDialogueID)
			End
		End If
		
		If g_sSide = "Right" And ((txtRightAnkle.Text = "" And txtRightLengths.Text = "") Or txtUidRightLeg.Text = "") Then
			MsgBox("No data for RIGHT Leg", 16, g_sDialogueID)
			End
		End If
		
		If g_sSide = "Right" And txtRightTemplate.Text = "" Then
			MsgBox("Figure RIGHT Leg before drawing", 16, g_sDialogueID)
			End
		End If
		
		
		
		'    If g_sSide = "Left" And fnGetNumber(txtLeftAnkle.Text, 1) = -1 And nLeftLegStyle <> 1 Then
		'        MsgBox "A Full leg has been chosen in the Body, but the data in the leg is footless.  ", 16, "First Leg Drawing"
		'        End
		'    End If
		
		
		'Open Macro file (fNum is declared as Global)
		sFile = "C:\JOBST\DRAW.D"
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = FN_WHOpen(sFile, txtPatientName, txtFileNo, g_sSide)
		
		BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WH_LEGC.D;")
		
		'Put standard patient details
		PR_StandardDetails()
		
		'Get Lengths and Circumferences
		Dim UseRevisedHts As Short
		'Heights
		If ARMEDDIA1.fnGetNumber(txtBody.Text, 4) > 0 Then UseRevisedHts = True Else UseRevisedHts = False
		If UseRevisedHts And ARMEDDIA1.fnGetNumber(txtBody.Text, 7) > 0 Then
			BDYUTILS.PR_PutStringAssign("sFoldHt", ARMEDDIA1.fnGetNumber(txtBody.Text, 7))
		Else
			BDYUTILS.PR_PutStringAssign("sFoldHt", txtFoldHt)
		End If
		
		If UseRevisedHts And ARMEDDIA1.fnGetNumber(txtBody.Text, 6) > 0 Then
			BDYUTILS.PR_PutStringAssign("sWaistHt", ARMEDDIA1.fnGetNumber(txtBody.Text, 6))
		Else
			BDYUTILS.PR_PutStringAssign("sWaistHt", txtWaistHt)
		End If
		
		If UseRevisedHts And ARMEDDIA1.fnGetNumber(txtBody.Text, 5) > 0 Then
			BDYUTILS.PR_PutStringAssign("sTOSHt", ARMEDDIA1.fnGetNumber(txtBody.Text, 5))
		Else
			BDYUTILS.PR_PutStringAssign("sTOSHt", txtTOSHt)
		End If
		
		
		'Body details
		BDYUTILS.PR_PutNumberAssign("nFoldHt", nFoldHt)
		BDYUTILS.PR_PutStringAssign("sWaistCir", txtWaistCir)
		BDYUTILS.PR_PutStringAssign("sTOSCir", txtTOSCir)
		BDYUTILS.PR_PutStringAssign("sTOSRed", txtTOSRed)
		BDYUTILS.PR_PutStringAssign("sWaistRed", txtWaistRed)
		BDYUTILS.PR_PutStringAssign("sFabric", txtFabric)
		BDYUTILS.PR_PutStringAssign("sLegStyle", txtLegStyle)
		
		BDYUTILS.PR_PutNumberAssign("nLegStyle", nLegStyle)
		BDYUTILS.PR_PutNumberAssign("nLeftLegStyle", nLeftLegStyle)
		BDYUTILS.PR_PutNumberAssign("nRightLegStyle", nRightLegStyle)
		
		'Leg Details
		If g_sSide = "Left" Then
			BDYUTILS.PR_PutLine("LeftLeg = %true;")
			BDYUTILS.PR_PutLine("RightLeg = %false;")
			sLengths = txtLeftLengths.Text 'Use later to establish 1st & Last tapes
			nValue = Val(Mid(sLengths, (6 * 4) + 1, 4)) / 10
			If nLeftLegStyle <> 0 Or nValue = 0 Then
				BDYUTILS.PR_PutLine("FootLess = %true;")
				Footless = True
			Else
				BDYUTILS.PR_PutLine("FootLess = %false;")
				'Ankle tape values
				BDYUTILS.PR_PutNumberAssign("nAnkleTape", ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 1))
				BDYUTILS.PR_PutStringAssign("sAnkleTape", Str(ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 1)))
				BDYUTILS.PR_PutStringAssign("sMMAnkle", Str(ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 2)))
				BDYUTILS.PR_PutStringAssign("sGramsAnkle", Str(ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 3)))
				BDYUTILS.PR_PutStringAssign("sReductionAnkle", Str(ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 4)))
				BDYUTILS.PR_PutNumberAssign("nReductionAnkle", ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 4))
			End If
			
			g_nFabricClass = ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 8)
			If g_nFabricClass = 2 Then
				BDYUTILS.PR_PutStringAssign("sReduction", txtLeftRed)
				BDYUTILS.PR_PutStringAssign("sTapeMMs", txtLeftMMs)
				BDYUTILS.PR_PutStringAssign("sStretch", txtLeftStr)
			End If
			
			BDYUTILS.PR_PutStringAssign("sFirstLeg", "Left")
			BDYUTILS.PR_PutStringAssign("sLeg", "Left")
			BDYUTILS.PR_PutStringAssign("sSecondLeg", "Right")
			
			g_sPressure = txtLeftTemplate.Text
			BDYUTILS.PR_PutStringAssign("sPressure", txtLeftTemplate)
			
			itemplate = Val(VB.Left(txtLeftTemplate.Text, 2))
			
			BDYUTILS.PR_PutStringAssign("sTapeLengths", txtLeftLengths)
			g_sTapeLength = txtLeftLengths.Text
			
			BDYUTILS.PR_PutStringAssign("sToeStyle", txtLeftToeStyle)
			
			g_nFootPleat1 = ARMEDDIA1.fnDisplaytoInches(Val(txtLeftFootPleat1.Text))
			BDYUTILS.PR_PutNumberAssign("nFootPleat1", g_nFootPleat1)
			
			g_nFootPleat2 = ARMEDDIA1.fnDisplaytoInches(Val(txtLeftFootPleat2.Text))
			BDYUTILS.PR_PutNumberAssign("nFootPleat2", g_nFootPleat2)
			
			g_nTopLegPleat1 = ARMEDDIA1.fnDisplaytoInches(Val(txtLeftTopLegPleat1.Text))
			BDYUTILS.PR_PutNumberAssign("nTopLegPleat1", g_nTopLegPleat1)
			
			g_nTopLegPleat2 = ARMEDDIA1.fnDisplaytoInches(Val(txtLeftTopLegPleat2.Text))
			BDYUTILS.PR_PutNumberAssign("nTopLegPleat2", g_nTopLegPleat2)
			
			BDYUTILS.PR_PutStringAssign("sFootLength", txtLeftFootLength)
			BDYUTILS.PR_PutNumberAssign("nElastic", ARMEDDIA1.fnGetNumber(txtLeftData.Text, 1))
			
			'Other leg details used to control toe
			nValue = ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 1)
			If nRightLegStyle <> 0 Or nValue < 0 Then
				BDYUTILS.PR_PutNumberAssign("nOtherAnkleMTapeLen ", 0)
			Else
				'Establish - Minus Ankle tape value
				If nValue = 8 Then
					'Ankle at +3
					nValue = Val(Mid(txtRightLengths.Text, (3 * 4) + 1, 4)) / 10
				Else
					'Ankle at +1-1/2
					nValue = Val(Mid(txtRightLengths.Text, (4 * 4) + 1, 4)) / 10
				End If
				BDYUTILS.PR_PutNumberAssign("nOtherAnkleMTapeLen ", ARMEDDIA1.fnDisplaytoInches(nValue))
			End If
			
		Else
			BDYUTILS.PR_PutLine("LeftLeg = %false;")
			BDYUTILS.PR_PutLine("RightLeg = %true;")
			sLengths = txtRightLengths.Text 'Use later to establish 1st & Last tapes
			nValue = Val(Mid(sLengths, (6 * 4) + 1, 4)) / 10
			If nRightLegStyle <> 0 Or nValue = 0 Then
				BDYUTILS.PR_PutLine("FootLess = %true;")
				Footless = True
			Else
				BDYUTILS.PR_PutLine("FootLess = %false;")
				'Ankle tape values
				BDYUTILS.PR_PutNumberAssign("nAnkleTape", ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 1))
				BDYUTILS.PR_PutStringAssign("sAnkleTape", Str(ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 1)))
				BDYUTILS.PR_PutStringAssign("sMMAnkle", Str(ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 2)))
				BDYUTILS.PR_PutStringAssign("sGramsAnkle", Str(ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 3)))
				BDYUTILS.PR_PutStringAssign("sReductionAnkle", Str(ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 4)))
				BDYUTILS.PR_PutNumberAssign("nReductionAnkle", ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 4))
			End If
			
			g_nFabricClass = ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 8)
			If g_nFabricClass = 2 Then
				BDYUTILS.PR_PutStringAssign("sReduction", txtRightRed)
				BDYUTILS.PR_PutStringAssign("sTapeMMs", txtRightMMs)
				BDYUTILS.PR_PutStringAssign("sStretch", txtRightStr)
			End If
			
			BDYUTILS.PR_PutStringAssign("sFirstLeg", "Right")
			BDYUTILS.PR_PutStringAssign("sLeg", "Right")
			BDYUTILS.PR_PutStringAssign("sSecondLeg", "Left")
			
			BDYUTILS.PR_PutStringAssign("sPressure", txtRightTemplate)
			g_sPressure = txtRightTemplate.Text
			itemplate = Val(VB.Left(txtRightTemplate.Text, 2))
			
			BDYUTILS.PR_PutStringAssign("sTapeLengths", txtRightLengths)
			g_sTapeLength = txtRightLengths.Text
			BDYUTILS.PR_PutStringAssign("sToeStyle", txtRightToeStyle)
			
			g_nFootPleat1 = ARMEDDIA1.fnDisplaytoInches(Val(txtRightFootPleat1.Text))
			g_nFootPleat2 = ARMEDDIA1.fnDisplaytoInches(Val(txtRightFootPleat2.Text))
			g_nTopLegPleat1 = ARMEDDIA1.fnDisplaytoInches(Val(txtRightTopLegPleat1.Text))
			g_nTopLegPleat2 = ARMEDDIA1.fnDisplaytoInches(Val(txtRightTopLegPleat2.Text))
			BDYUTILS.PR_PutNumberAssign("nFootPleat1", g_nFootPleat1)
			BDYUTILS.PR_PutNumberAssign("nFootPleat2", g_nFootPleat2)
			BDYUTILS.PR_PutNumberAssign("nTopLegPleat1", g_nTopLegPleat1)
			BDYUTILS.PR_PutNumberAssign("nTopLegPleat2", g_nTopLegPleat2)
			
			BDYUTILS.PR_PutStringAssign("sFootLength", txtRightFootLength)
			BDYUTILS.PR_PutNumberAssign("nElastic", ARMEDDIA1.fnGetNumber(txtRightData.Text, 1))
			
			'Other leg details used to control toe
			nValue = ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 1)
			If nLeftLegStyle <> 0 Or nValue < 0 Then
				BDYUTILS.PR_PutNumberAssign("nOtherAnkleMTapeLen ", 0)
			Else
				'Establish - Minus Ankle tape value
				If nValue = 8 Then
					'Ankle at +3
					nValue = Val(Mid(txtLeftLengths.Text, (3 * 4) + 1, 4)) / 10
				Else
					'Ankle at +1-1/2
					nValue = Val(Mid(txtLeftLengths.Text, (4 * 4) + 1, 4)) / 10
				End If
				BDYUTILS.PR_PutNumberAssign("nOtherAnkleMTapeLen ", ARMEDDIA1.fnDisplaytoInches(nValue))
			End If
		End If
		
		g_Footless = Footless
		
		'First and LastTapes
		g_nFirstTape = -1
		g_nLastTape = 30
		For ii = 0 To 29
			nValue = Val(Mid(sLengths, (ii * 4) + 1, 4)) / 10
			'Set first and last tape (assumes no holes in data)
			If g_nFirstTape < 0 And nValue > 0 Then g_nFirstTape = ii + 1
			If g_nLastTape = 30 And g_nFirstTape > 0 And nValue = 0 Then g_nLastTape = ii
		Next ii
		BDYUTILS.PR_PutNumberAssign("nFirstTape", g_nFirstTape)
		BDYUTILS.PR_PutNumberAssign("nLastTape", g_nLastTape)
		
		'Fabric Class (Load JOBSTEX_FL Procedures and Defaults if class = 2 )
		'For a footless style the only valid classes are 0 and 1
		If Footless And g_nFabricClass = 2 Then g_nFabricClass = 1
		BDYUTILS.PR_PutNumberAssign("nFabricClass", g_nFabricClass)
		
		
		
		'Draw leg
		'Load universal Procedures and Defaults
		BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHLG1DEF.D;")
		
		'Get Origin
		BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHLGCORG.D;")
		
		'Draw template
		BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHLGxTMP.D;")
		
		'Calculate Foot Points (If a foot exists)
		If Not Footless Then BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHFTPNTS.D;")
		
		'Draw Leg
		BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHLGCDWG.D;")
		
		'Close macro file
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
		
	End Sub
	
	Private Sub PR_WH_CutOut()
		'Procedure to draw a cutout for a waist height
		Dim sFile, sSide As String
		Dim sLengths As String
		Dim nLeftLegStyle, nLegStyle, nRightLegStyle As Short
		Dim nLastTape, nFirstTape, nFabricClass As Short
		Dim nUseRevisedHts, ii, Footless, nAge As Short
		Dim nWaistHt, nFoldHt, nTOSHt, nValue As Double
		Dim nOpenOff, nCuttingMarkOff, nHeelLength As Double
		Dim nBriefOff As Double
		
		'Establish Leg Style
		'and Open Macro file (fNum is declared as Global)
		nAge = Val(txtAge.Text)
		
		nLegStyle = ARMEDDIA1.fnGetNumber(txtLegStyle.Text, 1)
		sFile = "C:\JOBST\DRAW.D"
		
		Select Case nLegStyle
			Case 0, 1
				If ARMEDDIA1.fnDisplaytoInches(Val(txtLeftThighCir.Text)) - ARMEDDIA1.fnDisplaytoInches(Val(txtRightThighCir.Text)) > 1 Then
					sSide = "Right"
				Else
					sSide = "Left"
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				fNum = FN_WHOpen(sFile, txtPatientName, txtFileNo, sSide)
				BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WH_CUT.D;")
				BDYUTILS.PR_PutNumberAssign("nRightThighCir", ARMEDDIA1.fnDisplaytoInches(Val(txtRightThighCir.Text)))
				BDYUTILS.PR_PutNumberAssign("nLeftThighCir", ARMEDDIA1.fnDisplaytoInches(Val(txtLeftThighCir.Text)))
				
				If txtCrotchStyle.Text = "Open Crotch" Then
					BDYUTILS.PR_PutLine("OpenCrotch = %true;")
				Else
					BDYUTILS.PR_PutLine("OpenCrotch = %false;")
				End If
				
			Case 2
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				fNum = FN_WHOpen(sFile, txtPatientName, txtFileNo, sSide)
				sSide = "Left"
				BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WH_CUT.D;")
				BDYUTILS.PR_PutNumberAssign("nRightThighCir", ARMEDDIA1.fnDisplaytoInches(Val(txtRightThighCir.Text)))
				BDYUTILS.PR_PutNumberAssign("nLeftThighCir", ARMEDDIA1.fnDisplaytoInches(Val(txtLeftThighCir.Text)))
				
				If txtCrotchStyle.Text = "Open Crotch" Then
					BDYUTILS.PR_PutLine("OpenCrotch = %true;")
				Else
					BDYUTILS.PR_PutLine("OpenCrotch = %false;")
				End If
				
			Case 3
				sSide = "Left"
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				fNum = FN_WHOpen(sFile, txtPatientName, txtFileNo, sSide)
				BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WH_W4OC.D;")
				
				BDYUTILS.PR_PutLine("OpenCrotch = %true;")
				
				BDYUTILS.PR_PutNumberAssign("nLeftThighCir", ARMEDDIA1.fnDisplaytoInches(Val(txtLeftThighCir.Text)))
				BDYUTILS.PR_PutNumberAssign("nRightThighCir", ARMEDDIA1.fnDisplaytoInches(Val(txtLeftThighCir.Text)))
				
			Case 4
				sSide = "Right"
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				fNum = FN_WHOpen(sFile, txtPatientName, txtFileNo, sSide)
				BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WH_W4OC.D;")
				
				BDYUTILS.PR_PutLine("OpenCrotch = %true;")
				
				BDYUTILS.PR_PutNumberAssign("nRightThighCir", ARMEDDIA1.fnDisplaytoInches(Val(txtRightThighCir.Text)))
				BDYUTILS.PR_PutNumberAssign("nLeftThighCir", ARMEDDIA1.fnDisplaytoInches(Val(txtRightThighCir.Text)))
				
		End Select
		
		'Put standard patient details
		PR_StandardDetails()
		
		'Body details
		'Circumferences
		BDYUTILS.PR_PutNumberAssign("nTOSCir", ARMEDDIA1.fnDisplaytoInches(Val(txtTOSCir.Text)))
		BDYUTILS.PR_PutNumberAssign("nLargestCir", ARMEDDIA1.fnDisplaytoInches(Val(txtLargestCir.Text)))
		BDYUTILS.PR_PutNumberAssign("nMidPointCir", ARMEDDIA1.fnDisplaytoInches(Val(txtMidPointCir.Text)))
		BDYUTILS.PR_PutNumberAssign("nWaistCir", ARMEDDIA1.fnDisplaytoInches(Val(txtWaistCir.Text)))
		
		'Heights
		nUseRevisedHts = ARMEDDIA1.fnGetNumber(txtBody.Text, 4)
		nFoldHt = ARMEDDIA1.fnDisplaytoInches(Val(txtFoldHt.Text))
		If nUseRevisedHts > 0 And ARMEDDIA1.fnGetNumber(txtBody.Text, 7) > 0 Then
			nFoldHt = ARMEDDIA1.fnDisplaytoInches(ARMEDDIA1.fnGetNumber(txtBody.Text, 7))
		End If
		
		nWaistHt = ARMEDDIA1.fnDisplaytoInches(Val(txtWaistHt.Text))
		If nUseRevisedHts > 0 And ARMEDDIA1.fnGetNumber(txtBody.Text, 6) > 0 Then
			nWaistHt = ARMEDDIA1.fnDisplaytoInches(ARMEDDIA1.fnGetNumber(txtBody.Text, 6))
		End If
		
		nTOSHt = ARMEDDIA1.fnDisplaytoInches(Val(txtTOSHt.Text))
		If nUseRevisedHts > 0 And ARMEDDIA1.fnGetNumber(txtBody.Text, 5) > 0 Then
			nTOSHt = ARMEDDIA1.fnDisplaytoInches(ARMEDDIA1.fnGetNumber(txtBody.Text, 5))
		End If
		
		BDYUTILS.PR_PutNumberAssign("nFoldHt", nFoldHt)
		BDYUTILS.PR_PutNumberAssign("nTOSHt", nTOSHt)
		BDYUTILS.PR_PutNumberAssign("nWaistHt", nWaistHt)
		
		'Reductions
		BDYUTILS.PR_PutNumberAssign("nTOSGivenRed", Val(txtTOSRed.Text))
		BDYUTILS.PR_PutNumberAssign("nWaistGivenRed", Val(txtWaistRed.Text))
		BDYUTILS.PR_PutNumberAssign("nThighGivenRed", Val(txtThighRed.Text))
		BDYUTILS.PR_PutNumberAssign("nLargestGivenRed", Val(txtLargestRed.Text))
		BDYUTILS.PR_PutNumberAssign("nMidPointGivenRed", Val(txtMidPointRed.Text))
		
		'Open Crotch factors
		BDYUTILS.PR_PutNumberAssign("nCrotchFrontFactor", ARMEDDIA1.fnGetNumber(txtBody.Text, 1))
		BDYUTILS.PR_PutNumberAssign("nOpenFront", ARMEDDIA1.fnGetNumber(txtBody.Text, 2))
		BDYUTILS.PR_PutNumberAssign("nOpenBack", ARMEDDIA1.fnGetNumber(txtBody.Text, 3))
		
		
		BDYUTILS.PR_PutStringAssign("sCrotchStyle", txtCrotchStyle)
		
		nLegStyle = ARMEDDIA1.fnGetNumber(txtLegStyle.Text, 1)
		nLeftLegStyle = ARMEDDIA1.fnGetNumber(txtLegStyle.Text, 2)
		nRightLegStyle = ARMEDDIA1.fnGetNumber(txtLegStyle.Text, 3)
		
		BDYUTILS.PR_PutNumberAssign("nLegStyle", nLegStyle)
		BDYUTILS.PR_PutNumberAssign("nLeftLegStyle", nLeftLegStyle)
		BDYUTILS.PR_PutNumberAssign("nRightLegStyle", nRightLegStyle)
		
		
		'Leg Details
		If sSide = "Left" Then
			BDYUTILS.PR_PutLine("LeftLeg = %true;")
			BDYUTILS.PR_PutLine("RightLeg = %false;")
			sLengths = txtLeftLengths.Text 'Use later to establish 1st & Last tapes
			If nLeftLegStyle <> 0 Then
				BDYUTILS.PR_PutLine("FootLess = %true;")
				Footless = True
			Else
				BDYUTILS.PR_PutLine("FootLess = %false;")
			End If
			BDYUTILS.PR_PutStringAssign("sLeg", "Left")
		Else
			BDYUTILS.PR_PutLine("LeftLeg = %false;")
			BDYUTILS.PR_PutLine("RightLeg = %true;")
			sLengths = txtRightLengths.Text 'Use later to establish 1st & Last tapes
			If nRightLegStyle <> 0 Then
				BDYUTILS.PR_PutLine("FootLess = %true;")
				Footless = True
			Else
				BDYUTILS.PR_PutLine("FootLess = %false;")
			End If
			BDYUTILS.PR_PutStringAssign("sLeg", "Right")
		End If
		
		
		'Open crotch factors
		If txtCrotchStyle.Text = "Open Crotch" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object max(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			nHeelLength = ARMDIA1.max(ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 6), ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 6))
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
			BDYUTILS.PR_PutNumberAssign("nOpenOff", nOpenOff)
		End If
		
		'Brief offset
		If nLegStyle = 2 Then 'Brief
			nBriefOff = 1.25
			If txtCrotchStyle.Text = "Open Crotch" And nOpenOff = 0.375 Then nBriefOff = 1.625
			If txtCrotchStyle.Text = "Open Crotch" And nOpenOff = 0.75 Then nBriefOff = 2#
			If txtCrotchStyle.Text = "Horizontal Fly" Then nBriefOff = 1.625
			BDYUTILS.PR_PutNumberAssign("nBriefOff", nBriefOff)
		End If
		
		'Cutting marks for W4OC
		If nLegStyle = 3 Or nLegStyle = 4 Then 'W4OC-Left and W4OC-Right
			nCuttingMarkOff = 0.75
			If Footless = True Then 'ie panty
				If nAge <= 10 Then
					nCuttingMarkOff = 0.375
				End If
			Else
				If nHeelLength > 0 And nHeelLength < 9 Then
					nCuttingMarkOff = 0.375
				End If
			End If
			BDYUTILS.PR_PutNumberAssign("nCuttingMarkOff", nCuttingMarkOff)
		End If
		
		
		'First and LastTapes
		nFirstTape = -1
		nLastTape = 30
		For ii = 0 To 29
			nValue = Val(Mid(sLengths, (ii * 4) + 1, 4)) / 10
			'Set first and last tape (assumes no holes in data)
			If nFirstTape < 0 And nValue > 0 Then nFirstTape = ii + 1
			If nLastTape = 30 And nFirstTape > 0 And nValue = 0 Then nLastTape = ii
		Next ii
		BDYUTILS.PR_PutNumberAssign("nFirstTape", nFirstTape)
		BDYUTILS.PR_PutNumberAssign("nLastTape", nLastTape)
		
		
		'Load defaults and Procedures
		BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHCUTDEF.D;")
		
		'Draw cutout
		Select Case nLegStyle
			Case 0, 1
				BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHCUTDBD.D;")
			Case 2
				BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHCUTBRF.D;")
			Case 3, 4
				BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WH4OCDBD.D;")
				
		End Select
		
		'Close macro file
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		'The timer is enabled in Load.
		'A Timer event happens after 10 seconds have passed.
		'If this event happens we assume that there is problem and we kill
		'this instance
		'
		End
	End Sub
End Class
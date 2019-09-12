Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class stamp
	Inherits System.Windows.Forms.Form
	'   '* Windows API Functions Declarations
	'    Private Declare Function GetWindow Lib "User" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
	'    Private Declare Function GetWindowText Lib "User" (ByVal hwnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
	'    Private Declare Function GetWindowTextLength Lib "User" (ByVal hwnd As Integer) As Integer
	'    Private Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
	'    Private Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)
	'
	'   'Constanst used by GetWindow
	'    Const GW_CHILD = 5
	'    Const GW_HWNDFIRST = 0
	'    Const GW_HWNDLAST = 1
	'    Const GW_HWNDNEXT = 2
	'    Const GW_HWNDPREV = 3
	'    Const GW_OWNER = 4
	
	Private Sub Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Cancel.Click
		End
	End Sub
	
	Private Sub Edit_USER_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Edit_USER.Click
		EditUSER.ShowDialog()
		'Force click event to re-read values from file
		OptionUSER.Checked = False
		OptionUSER.Checked = True
	End Sub
	
	
	Private Function FN_TextCheck(ByRef sLine As Object) As String
		'Checks for the following DRAFIX special characters and escape
		'them using the backslash \
		'This function is required as the characters have a special
		'meaning to the DRAFIX interpreter.
		'
		'Check for :-
		'    (") - single double quote
		'    (\) - the backslash
		'
		Dim sNewline, sChar As String
		Dim QUOTE, BACKSLASH As String
		Dim ii As Short
		
		QUOTE = Chr(34)
		BACKSLASH = "\"
		
		For ii = 1 To Len(sLine)
			'UPGRADE_WARNING: Couldn't resolve default property of object sLine. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sChar = Mid(sLine, ii, 1)
			If sChar = BACKSLASH Or sChar = QUOTE Then
				sNewline = sNewline & BACKSLASH & sChar
			Else
				sNewline = sNewline & sChar
			End If
		Next ii
		
		FN_TextCheck = sNewline
		
	End Function
	
	'UPGRADE_ISSUE: Form event Form.LinkClose was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub Form_LinkClose()
		
		Select Case txtStampType.Text
			Case "Arm"
				OptionArm.Checked = True
			Case "Waist Ht"
				OptionWaistHt.Checked = True
			Case "Leg"
				OptionLeg.Checked = True
			Case "General"
				OptionGeneral.Checked = True
			Case "USER"
				OptionUSER.Checked = True
		End Select
		
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default ' Change pointer to hourglass.
		
		Show()
	End Sub
	
	Private Sub stamp_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Hide()
		'Check if a previous instance is running
		'If it is warn user and exit
		'UPGRADE_ISSUE: App property App.PrevInstance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		If App.PrevInstance Then
			MsgBox("The Stamp Module is already running!" & Chr(13) & "Use ALT-TAB and Cancel it.", 16, "Error Starting Stamps")
			End
		End If
		
		'Maintain while loading DDE data
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor ' Change pointer to hourglass.
		'Reset in Form_LinkClose
		
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.
		Edit_USER.Enabled = False
		
		MainForm = Me
		
	End Sub
	
	Private Sub OK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OK.Click
		prInsertStamp()
	End Sub
	
	'UPGRADE_WARNING: Event OptionArm.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub OptionArm_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptionArm.CheckedChanged
		If eventSender.Checked Then
			'Stamps for LEG
			Edit_USER.Enabled = False
			StampList.Items.Clear()
			StampList.Items.Add("1/2" & Chr(34) & " ELASTIC")
			StampList.Items.Add("2" & Chr(34) & " ELASTIC")
			StampList.Items.Add("SILICONE ELASTIC")
			StampList.Items.Add("NO ELASTIC")
			StampList.Items.Add("")
			StampList.Items.Add("LINING")
			StampList.Items.Add("INSIDE LINING")
			StampList.Items.Add("OUTSIDE LINING")
			StampList.Items.Add("FULL LINING")
			StampList.Items.Add("")
			StampList.Items.Add("REINFORCED ELBOW")
		End If
	End Sub
	
	'UPGRADE_WARNING: Event OptionGeneral.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub OptionGeneral_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptionGeneral.CheckedChanged
		If eventSender.Checked Then
			'Stamps for GENERAL use
			Edit_USER.Enabled = False
			StampList.Items.Clear()
			'StampList.AddItem "  "
		End If
	End Sub
	
	'UPGRADE_WARNING: Event OptionLeg.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub OptionLeg_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptionLeg.CheckedChanged
		If eventSender.Checked Then
			'Stamps for Legs
			Edit_USER.Enabled = False
			StampList.Items.Clear()
			StampList.Items.Add("REINFORCED KNEE")
			StampList.Items.Add("BEHIND KNEE LINING")
			StampList.Items.Add("REINFORCED TOP OF THIGH")
			StampList.Items.Add("NO ELASTIC")
			StampList.Items.Add("ELASTIC")
		End If
	End Sub
	
	'UPGRADE_WARNING: Event OptionUser.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub OptionUser_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptionUser.CheckedChanged
		If eventSender.Checked Then
			Dim hFileNumber As Object
			Dim sLine As Object
			
			'Stamps created and stored by the USER
			Edit_USER.Enabled = True
			StampList.Items.Clear()
			
			'Load USER Stamps from file
			'UPGRADE_WARNING: Couldn't resolve default property of object hFileNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			hFileNumber = FreeFile
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If Dir("C:\JOBST\STAMPS.USR") <> "" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object hFileNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileOpen(hFileNumber, "C:\JOBST\STAMPS.USR", OpenMode.Input)
				'UPGRADE_WARNING: Couldn't resolve default property of object hFileNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Do While Not EOF(hFileNumber)
					'UPGRADE_WARNING: Couldn't resolve default property of object hFileNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sLine = LineInput(hFileNumber)
					'UPGRADE_WARNING: Couldn't resolve default property of object sLine. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					StampList.Items.Add(sLine)
				Loop 
				'UPGRADE_WARNING: Couldn't resolve default property of object hFileNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileClose(hFileNumber)
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event OptionWaistHt.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub OptionWaistHt_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptionWaistHt.CheckedChanged
		If eventSender.Checked Then
			'Stamps for WAIST HT
			Edit_USER.Enabled = False
			StampList.Items.Clear()
			StampList.Items.Add("BEHIND KNEE LINING")
			StampList.Items.Add("REINFORCED KNEE")
			StampList.Items.Add("REINFORCED INNER\nTHIGH PERINEUM")
			StampList.Items.Add("REINFORCED TOP OF THIGH")
			StampList.Items.Add("")
			StampList.Items.Add("ATTACH SUSPENDERS")
			StampList.Items.Add("ATTACH VELCRO TABS\n& SEND OTHER PIECES")
			StampList.Items.Add("ATTACH VELCRO TABS\nTO VEST & WAIST HEIGHT")
			StampList.Items.Add("")
			StampList.Items.Add("2in ELASTIC WAIST BAND")
			StampList.Items.Add("DOUBLE WAIST BAND")
			StampList.Items.Add("")
			StampList.Items.Add("NO ELASTIC")
			StampList.Items.Add("ELASTIC")
			StampList.Items.Add("")
			StampList.Items.Add("CUT OUT \nON LEFT ONLY")
			
		End If
	End Sub
	
	Private Sub prInsertStamp()
		Dim nNewLineStart As Short
		Dim fNum As Short
		Dim sStampLine1, sStampLine2 As String
		Dim sPathJOBST As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object sPathJOBST. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sPathJOBST = fnPathJOBST()
		
		If txtStamp.Text <> "" Then
			'Split into two lines if required
			nNewLineStart = InStr(1, txtStamp.Text, "\n", 0)
			If nNewLineStart > 0 Then
				sStampLine1 = FN_TextCheck(VB.Left(txtStamp.Text, nNewLineStart - 1))
				sStampLine2 = FN_TextCheck(Mid(txtStamp.Text, nNewLineStart + 2))
			Else
				sStampLine1 = FN_TextCheck(txtStamp)
				sStampLine2 = ""
			End If
			
			'Create drawing macro
			'Open file
			fNum = FreeFile
			FileOpen(fNum, "C:\JOBST\DRAW.D", OpenMode.Output)
			
			'Write headder Information to file
			PrintLine(fNum, "//DRAFIX Macro created - " & DateString & "  " & TimeString)
			PrintLine(fNum, "//by Visual Basic STAMP.EXE")
			
			'Declare and set DRAFIX variables
			'N.B Chr$(34) is 'Double quotes (") character
			PrintLine(fNum, "STRING sStamp1, sStamp2;")
			PrintLine(fNum, "sStamp1=" & Chr(34) & sStampLine1 & Chr(34) & ";")
			PrintLine(fNum, "sStamp2=" & Chr(34) & sStampLine2 & Chr(34) & ";")
			
			'Select macro depending on Arrow check box
			'Note:-
			'    Two different macros are used to speed the insertion of text
			'    during the simple case with no arrows.
			If chkArrow.CheckState = 1 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object sPathJOBST. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "@" & sPathJOBST & "\stamps\stmparrw.d;")
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object sPathJOBST. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "@" & sPathJOBST & "\stamps\stmpdraw.d;")
			End If
			
			FileClose(fNum)
			
			'Activate DRAFIX Windows CAD
			AppActivate(fnGetDrafixWindowTitleText())
			
			'Start drawing macro
			System.Windows.Forms.SendKeys.SendWait("@c:\jobst\draw.d{enter}")
			
			End
		Else
			MsgBox("You have not yet selected a Stamp", 48, "Stamp Selection Error")
		End If
	End Sub
	
	'UPGRADE_WARNING: Event StampList.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub StampList_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles StampList.SelectedIndexChanged
		txtStamp.Text = VB6.GetItemString(StampList, StampList.SelectedIndex)
	End Sub
	
	Private Sub StampList_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles StampList.DoubleClick
		txtStamp.Text = VB6.GetItemString(StampList, StampList.SelectedIndex)
		If txtStamp.Text <> "" Then prInsertStamp()
	End Sub
End Class
Option Strict Off
Option Explicit On
Friend Class EditUSER
	Inherits System.Windows.Forms.Form
	'Index to stamp selected from StampList
	Dim nStampIndex As Short
	
	'Handle to file number
	Dim hFileNumber As Object
	
	
	Private Sub Add_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Add_Renamed.Click
		StampList.Items.Add(txtStamp.Text)
		'Deselect stamplist and disable buttons
		nStampIndex = -1
		StampList.SelectedIndex = nStampIndex
		Add_Renamed.Enabled = False
		Del.Enabled = False
		Change.Enabled = False
	End Sub
	
	Private Sub Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Cancel.Click
		Me.Close()
	End Sub
	
	Private Sub Change_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Change.Click
		StampList.Items.RemoveAt(nStampIndex)
		StampList.Items.Insert(nStampIndex, txtStamp.Text)
		'Deselect stamplist and disable buttons
		nStampIndex = -1
		StampList.SelectedIndex = nStampIndex
		Add_Renamed.Enabled = False
		Change.Enabled = False
		Del.Enabled = False
	End Sub
	
	Private Sub Del_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Del.Click
		StampList.Items.RemoveAt(nStampIndex)
		txtStamp.Text = ""
		'Deselect stamplist and disable buttons
		nStampIndex = -1
		StampList.SelectedIndex = nStampIndex
		Add_Renamed.Enabled = False
		Del.Enabled = False
		Change.Enabled = False
	End Sub
	
	Private Sub EditUSER_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim sLine As Object
		Dim sPathJOBST As String
		
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.
		
		'Read USER stamps from file
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
		
		'Disable Add, Delete and Change buttons
		Add_Renamed.Enabled = False
		Del.Enabled = False
		Change.Enabled = False
		nStampIndex = -1
		
	End Sub
	
	Private Sub OK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OK.Click
		Dim I As Object
		
		'Write revised USER stamps to file
		'UPGRADE_WARNING: Couldn't resolve default property of object hFileNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileOpen(hFileNumber, "C:\JOBST\STAMPS.USR", OpenMode.Output)
		For I = 0 To (StampList.Items.Count - 1)
			'UPGRADE_WARNING: Couldn't resolve default property of object hFileNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(hFileNumber, VB6.GetItemString(StampList, I))
		Next I
		'UPGRADE_WARNING: Couldn't resolve default property of object hFileNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(hFileNumber)
		
		'Return to main form
		Me.Close()
		
	End Sub
	
	'UPGRADE_WARNING: Event StampList.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub StampList_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles StampList.SelectedIndexChanged
		nStampIndex = StampList.SelectedIndex
		txtStamp.Text = VB6.GetItemString(StampList, nStampIndex)
		Del.Enabled = True
		Add_Renamed.Enabled = False
		Change.Enabled = False
	End Sub
	
	'UPGRADE_WARNING: Event txtStamp.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtStamp_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtStamp.TextChanged
		Add_Renamed.Enabled = True
		If nStampIndex >= 0 Then
			Change.Enabled = True
		End If
	End Sub
End Class
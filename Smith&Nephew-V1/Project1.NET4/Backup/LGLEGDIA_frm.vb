Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class lglegdia
	Inherits System.Windows.Forms.Form
	'Module :   LGLEGDIA.MAK
	'Purpose:   Input and Figure leg
	'
	'Version:   3.01
	'Date:      1995
	'Author:    Gary George
	'
	'Used in:
	'
	'-------------------------------------------------------
	'REVISIONS:
	'Date       By      Action
	'Oct/Nov 95 GG      Modifications w.r.t Triton / Imageable
	'
	'19.Dec.95  GG      Bug fix
	'                   Close button always assumed a change
	'                   had been made.
	'                   FN_ConcatData() added
	'
	'Jan 99     GG      Ported to VB5
	'-------------------------------------------------------
	'
	'Notes:-
	'    Much of the code and form has been hacked from
	'    WHFIGURE.FRM and WHLEGDIA.FRM
	'    As both of these are proven, there has been little
	'    to no changes made except to disable JOBSTEX_FL
	'
	'    Reference to the left leg should be taken to mean
	'    the current leg
	'
	
	'MsgBox constants
	'    Const IDYES = 6
	'    Const IDNO = 7
	'    Const IDCANCEL = 2
	
	'Other constants
	Const NEWLINE As Short = 13
	
	
	Private Sub Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Cancel.Click
		Dim Response As Short
		Dim sTask As String
		
		'Check if data has been modified
		If g_sChangeChecker <> FN_ConcatData() Then
			Response = MsgBox("Changes have been made, Save changes before closing", 35, "LEG Details Dialogue")
			Select Case Response
				Case IDYES
					Update_DDE_Text_Boxes()
					'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					fnum = FN_SaveOpen("c:\jobst\draw.d", txtPatientName, txtFileNo, txtLeg)
					PR_SaveLeg()
					'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					FileClose(fnum)
					sTask = fnGetDrafixWindowTitleText()
					If sTask <> "" Then
						AppActivate(fnGetDrafixWindowTitleText())
						System.Windows.Forms.SendKeys.SendWait("@c:\jobst\draw.d{enter}")
						End
					Else
						MsgBox("Can't find a Drafix Drawing to update!", 16, "LEG Details Dialogue")
					End If
				Case IDNO
					End
				Case IDCANCEL
					Exit Sub
			End Select
		Else
			End
		End If
		
	End Sub
	
	'UPGRADE_WARNING: Event cboFabric.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboFabric_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboFabric.SelectedIndexChanged
		PR_FigureLeftAnkle()
	End Sub
	
	'UPGRADE_WARNING: Event chkLeftZipper.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkLeftZipper_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLeftZipper.CheckStateChanged
		If g_JOBSTEX = True Or g_JOBSTEX_FL = True Then
			If chkLeftZipper.CheckState = 1 Then
				cboLeftTemplate.SelectedIndex = 1 '9DS
			Else
				cboLeftTemplate.SelectedIndex = 0 '13DS
			End If
		End If
		PR_FigureLeftAnkle()
	End Sub
	
	Private Sub Draw_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Draw.Click
		Dim sCAD_App As String
		'Check that data is all present and insert into drafix
		If Validate_Data() Then
			Update_DDE_Text_Boxes()
			PR_DrawAndSaveLeg()
			sCAD_App = fnGetDrafixWindowTitleText()
			If sCAD_App <> "" Then
				AppActivate(sCAD_App)
				System.Windows.Forms.SendKeys.SendWait("@C:\JOBST\DRAW.D{enter}")
				End
			Else
				MsgBox("Can't find drafix", 16, "Leg Dialogue")
			End If
		End If
	End Sub
	
	Private Sub ExtendLegTapes_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ExtendLegTapes.Click
		'Takes the last tape given and add the required value
		'GOP - 01-02/12, 6.2.1
		'
		Dim ii As Short
		Dim nValue As Double
		
		'Locate last tape
		For ii = 29 To 0 Step -1
			If Val(txtLeft(ii).Text) > 0 Then Exit For
		Next ii
		
		'Check that there are some tapes.
		'Check that there is room to add a new tape
		If ii = 0 Or ii = 29 Then
			Beep()
			Exit Sub
		End If
		
		'Convert given value to inches
		nValue = ARMEDDIA1.fnDisplaytoInches(Val(txtLeft(ii).Text))
		
		'Extend tape
		If txtSex.Text = "Male" Then
			nValue = nValue + 0.5
		Else
			nValue = nValue + 1
		End If
		
		If g_iStyleLastTape = ii Then PR_SetLastTape(ii + 1)
		
		txtLeft(ii + 1).Text = CStr(BDLEGDIA1.fnInchesToDisplay(nValue))
		grdLeftInches.Row = ii + 1
		grdLeftInches.Text = ARMEDDIA1.fnInchesToText(nValue)
		
	End Sub
	
	Private Function FN_BuildStyleString() As String
		
		Dim sString, sFabricClass As String
		Dim sFootLength As String
		Dim nSave As Short
		
		If g_POWERNET = True Then sFabricClass = "0"
		If g_JOBSTEX = True Then sFabricClass = "1"
		If g_JOBSTEX_FL = True Then sFabricClass = "2"
		
		nSave = 0
		
		'Transfer to the "legbox" the following set of data in blank delimited format
		'these will be stored in the fields relevent data fields
		'
		'                            DATA - Field
		'                            ~~~~~~~~~~~~
		'                LegStyle                            (1)
		'                First Tape of style                 (2)
		'                Last Tape of style                  (3)
		'                AnkleTape#                          (4)
		'                Pressure                            (5)
		'                [Grams|Stretch]                     (6)
		'                Reduction                           (7)
		'                AnkleLength                         (8)
		'                HeelLength                          (9)
		'                Zipper Status                       (10)
		'                FabricClass                         (11)
		'                Toe Style %%%                       (12)
		'                Foot Length                         (13)
		'                Template                            (14)
		'                Heel Style  ***                     (15)
		'                Heel Reinforcement ***              (16)
		'
		'          *** = Not yet implemented
		'          %%% = Above Knee / Below Knee for Footless styles
		'
		'Where Leg Style has the following meanings :-
		'
		'    0 = Anklet
		'    1 = Knee High
		'    2 = Thigh High
		'    3 = Knee band
		'    4 = Thigh band
		'
		'The First and Last tape positions are style dependant.
		If txtFootLength.Text = "" Then
			sFootLength = "0"
		Else
			sFootLength = Str(CDbl(txtFootLength.Text))
		End If
		
		sString = Str(g_iLegStyle) & " "
		If g_iLegStyle < 3 Then
			sString = sString & Str(g_iFirstTape + nSave) & " "
		Else
			sString = sString & Str(g_iStyleFirstTape + nSave) & " "
		End If
		sString = sString & Str(g_iStyleLastTape + nSave) & " "
		
		Select Case g_iLegStyle
			Case 0
				If g_iLtAnkle <> 0 Then
					sString = sString & Str(g_iLtAnkle + 1) & " "
					sString = sString & "0 "
					sString = sString & "0 "
					sString = sString & "0 "
					sString = sString & Str(g_nLtLastAnkle) & " "
					sString = sString & Str(g_nLtLastHeel) & " "
					sString = sString & Str(g_iLtLastZipper) & " "
					sString = sString & sFabricClass & " "
					sString = sString & Str(cboToeStyle.SelectedIndex) & " "
					sString = sString & sFootLength & " "
					sString = sString & Str(cboLeftTemplate.SelectedIndex) & " -1 -1"
				Else
					'Set to reflect the footless state
					sString = sString & "-1 0 0 0 0 0 0 "
					sString = sString & sFabricClass
					sString = sString & " 0 0 "
					sString = sString & Str(cboLeftTemplate.SelectedIndex) & " -1 -1"
				End If
			Case 1, 2
				If g_iLtAnkle <> 0 Then
					sString = sString & Str(g_iLtAnkle + 1) & " "
					sString = sString & Str(g_iLtMM(g_iLtAnkle)) & " "
					sString = sString & Str(g_iLtStretch(g_iLtAnkle)) & " "
					sString = sString & Str(g_iLtRed(g_iLtAnkle)) & " "
					sString = sString & Str(g_nLtLastAnkle) & " "
					sString = sString & Str(g_nLtLastHeel) & " "
					sString = sString & Str(g_iLtLastZipper) & " "
					sString = sString & sFabricClass & " "
					sString = sString & Str(cboToeStyle.SelectedIndex) & " "
					sString = sString & sFootLength & " "
					sString = sString & Str(cboLeftTemplate.SelectedIndex) & " -1 -1"
				Else
					'Set to reflect the footless state
					sString = sString & "-1 0 0 0 0 0 0 "
					sString = sString & sFabricClass
					sString = sString & " 0 0 "
					sString = sString & Str(cboLeftTemplate.SelectedIndex) & " -1 -1"
				End If
				
			Case 3, 4, 5
				sString = sString & "-1 0 0 0 0 0 0 "
				sString = sString & sFabricClass
				sString = sString & " 0 0 "
				sString = sString & Str(cboLeftTemplate.SelectedIndex) & " -1 -1"
		End Select
		
		FN_BuildStyleString = sString
		
	End Function
	
	Private Function FN_ConcatData() As String
		'Concatenates the displayed data
		'this can then be used to check if any modifications have been made
		'
		Dim ii As Short
		Dim vData As Object
		
		'Initialise to blank string
		'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		vData = ""
		
		For ii = 0 To 29
			'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			vData = vData + txtLeft(ii).Text
			'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If ii >= 6 Then vData = vData + txtLeftMM(ii).Text
		Next ii
		For ii = 0 To 1
			'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			vData = vData + Str(optFabric(ii).Checked)
		Next ii
		
		For ii = 0 To 5
			'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			vData = vData + Str(optType(ii).Checked)
		Next ii
		
		'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		vData = vData + cboFabric.Text + cboToeStyle.Text + txtFootLength.Text + txtFirstTape.Text + txtLastTape.Text + cboLeftTemplate.Text + Str(chkLeftZipper.CheckState)
		'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		vData = vData + txtFootPleat1.Text + txtFootPleat2.Text + txtTopLegPleat1.Text + txtTopLegPleat1.Text
		
		'UPGRADE_WARNING: Couldn't resolve default property of object vData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FN_ConcatData = vData
		
	End Function
	
	Private Function FN_DrawOpen(ByRef sDrafixFile As String, ByRef sName As Object, ByRef sPatientFile As Object, ByRef sLeftorRight As Object) As Short
		'Open the DRAFIX macro file
		'Initialise Global variables
		
		'Open file
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fnum = FreeFile
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileOpen(fnum, sDrafixFile, OpenMode.Output)
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FN_DrawOpen = fnum
		
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
		
		'Initialise patient globals
		'UPGRADE_WARNING: Couldn't resolve default property of object sPatientFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_sFileNo = sPatientFile
		'UPGRADE_WARNING: Couldn't resolve default property of object sLeftorRight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_sSide = sLeftorRight
		'UPGRADE_WARNING: Couldn't resolve default property of object sName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_sPatient = sName
		
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
		
		'Write header information etc. to the DRAFIX macro file
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "//DRAFIX Leg Drawing Macro created - " & DateString & "  " & TimeString)
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "//Patient - " & g_sPatient & CC & " " & g_sFileNo & CC & " SIDE - " & g_sSide)
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "//by Visual Basic")
		
		'Text data
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "SetData(" & QQ & "TextHorzJust" & QC & g_nCurrTextHorizJust & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "SetData(" & QQ & "TextVertJust" & QC & g_nCurrTextVertJust & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "SetData(" & QQ & "TextHeight" & QC & g_nCurrTextHt & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "SetData(" & QQ & "TextAspect" & QC & g_nCurrTextAspect & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "SetData(" & QQ & "TextFont" & QC & g_nCurrTextFont & ");")
		
		BDYUTILS.PR_PutLine("STRING sTitleName, sLeg, sTmp, sWorkOrder, sID, sPathJOBST;")
		
		'Path to JOBST installed directory
		BDYUTILS.PR_PutStringAssign("sPathJOBST", ARMDIA1.FN_EscapeSlashesInString(g_sPathJOBST))
		
	End Function
	
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
	
	
	Private Function FN_POWERNET_Pressure(ByRef iReduction As Short, ByRef nAnkleCir As Double, ByRef iModulus As Short) As Short
		'nMMHg = FN_POWERNET_Pressure(iReduction, nAnkleCir, iModulus)
		'This fuction is the reverse of FN_POWERNET_reduction
		'
		'Input
		'    iReduction  a Reduction in the range 10 to 32
		'    iModulus    Modulus of chosen fabric
		'                Eg Fabric = "Pow 210-3B Cream" => Modulus = 210.
		'    nAnkleCir   Ankle circumferance used with the derved grams to
		'                back calculat the pressue at the given reductio
		'
		'Globals
		'    POWERNET    The Fabric conversion chart loaded from file.
		'                Maps modulus to reduction at the given grams.
		'
		'Output
		'    nMMHg       Pressure reverse calclated from the conversion chart.
		'
		'
		'NOTE:-
		'    No range checking is done on the reduction values.
		'    Similarly the modulus is not checked.
		'    So it had better be right.
		'
		Dim sConversion As String
		Dim iGrams, iVal, ii As Short
		
		'Get conversion string based on modulus
		sConversion = ""
		For ii = 0 To 17
			If iModulus = Val(POWERNET.Modulus(ii)) Then
				sConversion = POWERNET.Conversion_Renamed(ii)
				Exit For
			End If
		Next ii
		
		iVal = iReduction - 10
		iGrams = Val(Mid(sConversion, (iVal * 4) + 1, 4))
		
		FN_POWERNET_Pressure = ARMDIA1.round(iGrams / nAnkleCir)
		
	End Function
	
	Private Function FN_POWERNET_Reduction(ByRef iGrams As Short, ByRef iModulus As Short) As Short
		'nReduction = FN_POWERNET_Reduction(nMMHg, nGrams, iModulus)
		'Input
		'    iGrams      Grams at Ankle
		'    iModulus    Modulus of chosen fabric
		'                Eg Fabric = "Pow 210-3B Cream" => Modulus = 210.
		'
		'Globals
		'    POWERNET    The Fabric conversion chart loaded from file.
		'                Maps modulus to reduction at the given grams.
		'
		'Output
		'    iReduction  Reduction established from the conversion chart.
		'                In the range 10 to 32
		'
		'
		'NOTE
		'    This fuction is derived from the DRAFIX function
		'    FNCalcReduction()
		'
		Dim sConversion As String
		Dim iPrevVal, ii, iVal As Short
		
		'Get conversion string based on modulus
		sConversion = ""
		For ii = 0 To 17
			If iModulus = Val(POWERNET.Modulus(ii)) Then
				sConversion = POWERNET.Conversion_Renamed(ii)
				Exit For
			End If
		Next ii
		
		If sConversion = "" Then
			FN_POWERNET_Reduction = -1000
			Exit Function
		End If
		
		iPrevVal = 0
		For ii = 0 To 22
			iVal = Val(Mid(sConversion, (ii * 4) + 1, 4))
			If iVal >= iGrams Then Exit For
			iPrevVal = iVal
		Next ii
		
		'Default to a 10 reduction
		If ii = 0 Then
			FN_POWERNET_Reduction = 10
			Exit Function
		End If
		
		'Get reduction closest to given grams
		If (iGrams - iPrevVal) < (iVal - iGrams) Then
			FN_POWERNET_Reduction = ii + 9
		Else
			FN_POWERNET_Reduction = ii + 10
		End If
		
	End Function
	
	Private Function FN_SaveOpen(ByRef sDrafixFile As String, ByRef sName As Object, ByRef sPatientFile As Object, ByRef sLeftorRight As Object) As Short
		
		'Open the DRAFIX macro file
		'Initialise Global variables
		
		'Open file
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fnum = FreeFile
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileOpen(fnum, sDrafixFile, OpenMode.Output)
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FN_SaveOpen = fnum
		
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
		
		'Initialise patient globals
		'UPGRADE_WARNING: Couldn't resolve default property of object sPatientFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_sFileNo = sPatientFile
		'UPGRADE_WARNING: Couldn't resolve default property of object sLeftorRight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_sSide = sLeftorRight
		'UPGRADE_WARNING: Couldn't resolve default property of object sName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_sPatient = sName
		
		'Write header information etc. to the DRAFIX macro file
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "//DRAFIX Leg Drawing Macro created - " & DateString & "  " & TimeString)
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "//Patient - " & g_sPatient & CC & " " & g_sFileNo & CC & " SIDE - " & g_sSide)
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "//by Visual Basic")
		BDYUTILS.PR_PutLine("STRING sTitleName, sLeg, sTmp, sWorkOrder, sID, sPathJOBST;")
		
		'Path to JOBST installed directory
		BDYUTILS.PR_PutStringAssign("sPathJOBST", ARMDIA1.FN_EscapeSlashesInString(g_sPathJOBST))
		
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
		Dim sMessage As String
		Dim iLastStyleFound As Short
		Dim sSpacer As Object
		
		'Disable timeout timer
		Timer1.Enabled = False
		
		'Units
		If txtUnits.Text = "cm" Then
			g_nUnitsFac = 10 / 25.4
		Else
			g_nUnitsFac = 1
		End If
		
		'Use Title fabric if a fabric is not already given
		If txtLegTitleFabric.Text <> "" And txtFabric.Text = "" Then txtFabric.Text = txtLegTitleFabric.Text
		
		'Set dropdown combo boxes
		'Toe Style
		cboToeStyle.Items.Add("")
		cboToeStyle.Items.Add("Curved")
		cboToeStyle.Items.Add("Cut-Back")
		cboToeStyle.Items.Add("Straight")
		cboToeStyle.Items.Add("Soft Enclosed")
		cboToeStyle.Items.Add("Soft Enclosed B/M")
		cboToeStyle.Items.Add("Self Enclosed")
		
		'Set value
		For ii = 0 To (cboToeStyle.Items.Count - 1)
			If VB6.GetItemString(cboToeStyle, ii) = txtToeStyle.Text Then
				cboToeStyle.SelectedIndex = ii
			End If
		Next ii
		
		'Leg Option buttons
		If txtLeg.Text = "Left" Then
			Me.Text = "LEG Details - Left"
			optLeftLeg.Checked = True
			optRightLeg.Enabled = False
		End If
		If txtLeg.Text = "Right" Then
			Me.Text = "LEG Details - Right"
			optRightLeg.Checked = True
			optLeftLeg.Enabled = False
		End If
		
		'Update tape boxes
		'Get first and last tapes in the text boxes
		g_iFirstTape = -1
		g_iLastTape = -1
		
		g_iStyleFirstTape = -1
		g_iStyleLastTape = 30
		
		nValue = 0
		grdLeftInches.Col = 0
		For ii = 0 To 29
			nValue = Val(Mid(txtLeftLengths.Text, (ii * 4) + 1, 4)) / 10
			If nValue > 0 Then
				txtLeft(ii).Text = CStr(nValue)
				nValue = ARMEDDIA1.fnDisplaytoInches(nValue)
				grdLeftInches.Row = ii
				grdLeftInches.Text = ARMEDDIA1.fnInchesToText(nValue)
			End If
			If g_iFirstTape < 0 And nValue > 0 Then g_iFirstTape = ii
			If g_iLastTape < 0 And g_iFirstTape > 0 And nValue = 0 Then g_iLastTape = ii - 1
		Next ii
		If nValue > 0 Then g_iLastTape = 29
		
		'Set leg style options
		'The chosen leg style is given when the user has selected a
		'pattern from the drawing rather than using one of the buttons.
		
		g_iLegStyle = ARMEDDIA1.fnGetNumber(txtChosenStyle.Text, 1)
		
		If g_iLegStyle >= 0 Then
			optType_CheckedChanged(optType.Item(g_iLegStyle), New System.EventArgs())
			optType(g_iLegStyle).Checked = True
			PR_SetLastTape(g_iStyleLastTape)
		Else
			PR_LastTapeDisplay("Disabled")
			PR_FirstTapeDisplay("Disabled")
		End If
		
		If g_iLegStyle >= 3 Then PR_SetFirstTape(g_iStyleFirstTape)
		
		
		'Establish how the leg is to be figured
		'Base this on chosen fabric and availability of mm
		'If it is not given explicitly
		
		'Check for an explititly given fabric class from the Style field
		'This Field will be emplty if no previous figuring has been done
		'Note This is a multi data. I the first number is -ve then there is no ankle.
		'     Also the function fnGetNumber returns -1 if the number does not exist.
		'
		iFabricClass = -1
		If ARMEDDIA1.fnGetNumber(g_sStyleString, 4) <> -1 Then
			iFabricClass = ARMEDDIA1.fnGetNumber(g_sStyleString, 11)
		End If
		
		'Get fabric class from other data
		If iFabricClass < 0 Then
			If txtFabric.Text = "" Then
				If txtDiagnosis.Text = "Burns" Then
					g_POWERNET = True
				Else
					'                If Left$(txtDiagnosis.Text, 5) = "Lymph" Then
					'                    g_JOBSTEX_FL = True
					'                Else
					'                    g_JOBSTEX = True
					'                End If
					g_JOBSTEX = True
				End If
			Else
				If VB.Left(txtFabric.Text, 3) = "Pow" Then
					g_POWERNET = True
				Else
					'                If txtLeftMMs.Text <> Then
					'                    g_JOBSTEX_FL = True
					'                Else
					'                    g_JOBSTEX = True
					'                End If
					g_JOBSTEX = True
				End If
			End If
		Else
			If iFabricClass = 0 Then g_POWERNET = True
			If iFabricClass = 1 Then g_JOBSTEX = True
			If iFabricClass = 2 Then g_JOBSTEX = True
			'        If iFabricClass = 2 Then g_JOBSTEX_FL = True
		End If
		
		If g_JOBSTEX = True Then optFabric(1).Checked = True
		'   If g_JOBSTEX_FL = True Then optFabric(2).Value = True
		
		'Setup fabric dropdown box
		If g_POWERNET = True Then
			optFabric(0).Checked = True
			LGLEGDIA1.PR_GetComboListFromFile(cboFabric, g_sPathJOBST & "\WHFABRIC.DAT")
			ARMEDDIA1.PR_LoadFabricFromFile(g_sPathJOBST & "\TEMPLTS\POWERNET.DAT")
		End If
		
		If g_JOBSTEX = True Or g_JOBSTEX_FL = True Then
			'Jobstex fabric
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
		End If
		
		'Set fabric value
		'NB   It might be that the fabric originally used is not one on the
		'     current fabric list.  In this case add the given fabric to
		'     the start of the list
		For ii = 0 To (cboFabric.Items.Count - 1)
			If VB6.GetItemString(cboFabric, ii) = txtFabric.Text Then
				cboFabric.SelectedIndex = ii
			End If
		Next ii
		If txtFabric.Text <> "" And cboFabric.SelectedIndex = -1 Then
			cboFabric.Items.Insert(0, txtFabric.Text)
			cboFabric.SelectedIndex = 0
		End If
		
		'Set up depending on leg style chosen
		'if only a single style available then display the values for that
		'style only
		'other wise display as blank untill the user selects an option
		
		PR_EstablishAnkles()
		
		If g_POWERNET = True Then PR_EnablePOWERNET()
		If g_JOBSTEX = True Or g_JOBSTEX_FL = True Then PR_EnableJOBSTEX()
		'   If g_JOBSTEX_FL = True Then PR_EnableFL_JOBSTEX
		
		
		'From the g_sStyleString set above extract the saved ankle figuring
		'We have to be careful in the order in which this is set up as in the case
		'of JOBSTEX_FL we do not want to cause recaculation on the ankle.
		'We also must take care only to add values only if they exist
		'(thus we check that the returned value from fnGetNumber is not -ve)
		'
		Dim iFabric As Short
		If g_sStyleString <> "" And g_iLtAnkle <> 0 Then
			iZipper = ARMEDDIA1.fnGetNumber(g_sStyleString, 10)
			If iZipper >= 0 Then chkLeftZipper.CheckState = iZipper 'Order is v.important here
			iMMHg = ARMEDDIA1.fnGetNumber(g_sStyleString, 5)
			If iMMHg >= 0 Then
				g_iLtMM(g_iLtAnkle) = iMMHg
				txtLeftMM(g_iLtAnkle).Text = CStr(iMMHg)
			End If
			If g_JOBSTEX_FL = True Then
				nStretch = ARMEDDIA1.fnGetNumber(g_sStyleString, 6)
				iFabric = Val(VB.Left(cboFabric.Text, 2))
				If nStretch >= 0 Then
					PR_DisplayFiguredAnkle("Left", nStretch, g_nLtLastAnkle, g_nLtLastHeel, iFabric)
					g_iLtStretch(g_iLtAnkle) = nStretch
				End If
			End If
		End If
		
		
		'Don't figure ankle for fabric class JOBSTEX_FL as the user can modify the pressures
		'at each leg tape manually and these are stored.  Figuring would overwrite
		'any manual changes
		If g_JOBSTEX_FL <> True Then
			PR_FigureLeftAnkle()
		Else
			If txtLeftMMs.Text <> "" And g_iLtAnkle <> 0 Then
				For ii = g_iLtAnkle + 1 To 29
					iValue = Val(Mid(txtLeftMMs.Text, (ii * 3) + 1, 3))
					If iValue > 0 Then
						txtLeftMM(ii).Text = CStr(iValue)
						PR_FigureLeftTape(ii)
					End If
				Next ii
			End If
		End If
		
		'Disable anklet and knee if no ankle tape
		If g_iLtAnkle = 0 Then
			optType(0).Enabled = False
			optType(1).Enabled = False
			optType(2).Enabled = False
		End If
		
		'Find existing styles display last found
		sMessage = ""
		'UPGRADE_WARNING: Couldn't resolve default property of object sSpacer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sSpacer = ""
		iLastStyleFound = -1
		
		If txtAnklet.Text <> "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object sSpacer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sMessage = sMessage + sSpacer + "Anklet"
			'UPGRADE_WARNING: Couldn't resolve default property of object sSpacer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sSpacer = ", "
			iLastStyleFound = 0
		End If
		
		If txtKneeLength.Text <> "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object sSpacer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sMessage = sMessage + sSpacer + "Knee Length"
			'UPGRADE_WARNING: Couldn't resolve default property of object sSpacer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sSpacer = ", "
			iLastStyleFound = 1
		End If
		If txtThighLength.Text <> "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object sSpacer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sMessage = sMessage + sSpacer + "Thigh Length"
			'UPGRADE_WARNING: Couldn't resolve default property of object sSpacer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sSpacer = ", "
			iLastStyleFound = 2
		End If
		If txtKneeBand.Text <> "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object sSpacer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sMessage = sMessage + sSpacer + "Knee Band"
			'UPGRADE_WARNING: Couldn't resolve default property of object sSpacer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sSpacer = ", "
			iLastStyleFound = 3
		End If
		If txtThighBandAK.Text <> "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object sSpacer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sMessage = sMessage + sSpacer + "Thigh (A/K)"
			'UPGRADE_WARNING: Couldn't resolve default property of object sSpacer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sSpacer = ", "
			iLastStyleFound = 4
		End If
		If txtThighBandBK.Text <> "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object sSpacer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sMessage = sMessage + sSpacer + "Thigh (B/K)"
			'UPGRADE_WARNING: Couldn't resolve default property of object sSpacer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sSpacer = ", "
			iLastStyleFound = 5
		End If
		
		If iLastStyleFound >= 0 Then
			optType_CheckedChanged(optType.Item(iLastStyleFound), New System.EventArgs())
			optType(iLastStyleFound).Checked = True
			labMessage.Text = "Existing Styles: " & Chr(13) & sMessage
		End If
		
		g_sChangeChecker = FN_ConcatData()
		Show()
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default 'Change pointer to default.
		Me.Activate()
		
	End Sub
	
	'UPGRADE_ISSUE: Form event Form.LinkExecute was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub Form_LinkExecute(ByRef CmdStr As String, ByRef Cancel As Short)
		If CmdStr = "Cancel" Then
			Cancel = 0
			End
		End If
	End Sub
	
	Private Sub lglegdia_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim ii As Short
		Hide()
		'Check if a previous instance is running
		'If it is warn user and exit
		'UPGRADE_ISSUE: App property App.PrevInstance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		If App.PrevInstance Then
			MsgBox("The Leg input Module is already running!" & Chr(13) & "Use ALT-TAB and Cancel it.", 16, "Error Starting Figure")
			End
		End If
		
		'Maintain while loading DDE data
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor ' Change pointer to hourglass.
		'Reset in Form_LinkClose
		
		'Position to center of screen
		Left_Renamed.Text = CStr((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2) ' Center form horizontally.
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2) ' Center form vertically.
		
		MainForm = Me
		
		g_nUnitsFac = 1 'Default to inches
		g_iLegStyle = -1
		g_JOBSTEX = False
		g_JOBSTEX_FL = False
		g_POWERNET = False
		g_sTextList = "-7½ -6-4½ -3-1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"
		
		'Setup display inches grid
		grdLeftInches.set_ColWidth(0, 880)
		grdLeftInches.set_ColAlignment(0, 2)
		
		For ii = 0 To 29
			grdLeftInches.set_RowHeight(ii, 266)
		Next ii
		
		'Setup display of results grid
		For ii = 0 To 1
			grdLeftDisplay.set_ColWidth(ii, 488)
			grdLeftDisplay.set_ColAlignment(ii, 2)
		Next ii
		
		For ii = 0 To 23
			grdLeftDisplay.set_RowHeight(ii, 266)
		Next ii
		
		
		g_sPathJOBST = fnPathJOBST()
		
		'Enable time out timer
		Timer1.Interval = 6000
		Timer1.Enabled = True
		
	End Sub
	
	'UPGRADE_WARNING: Event optFabric.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optFabric_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optFabric.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optFabric.GetIndex(eventSender)
			'Allows the selection of fabric class
			'Used to override the fabric class selection established at link close
			
			'Clean Message box
			labMessage.Text = ""
			
			'Do nothing if same fabric class selected
			If g_POWERNET = True And Index = 0 Then Exit Sub
			If g_JOBSTEX = True And Index = 1 Then Exit Sub
			If g_JOBSTEX_FL = True And Index = 2 Then Exit Sub
			
			'Restablish ankles just in case they have changed
			PR_EstablishAnkles()
			
			'Set the display for the different fabric classes
			'We must first reset the form to the blank form that we use on form load
			PR_ResetFormLGLEGDIA()
			
			g_POWERNET = False
			g_JOBSTEX = False
			g_JOBSTEX_FL = False
			
			Select Case Index
				Case 0
					'POWERNET
					g_POWERNET = True
					PR_EnablePOWERNET()
				Case 1
					'JOBSTEX
					g_JOBSTEX = True
					PR_EnableJOBSTEX()
				Case 2
					'JOBSTEX_FL, Pressure calulated at every leg tape
					g_JOBSTEX_FL = True
					PR_EnableJOBSTEX()
					'PR_EnableFL_JOBSTEX
			End Select
			
			'Setup fabric dropdown box
			cboFabric.Items.Clear()
			If g_POWERNET = True Then
				LGLEGDIA1.PR_GetComboListFromFile(cboFabric, g_sPathJOBST & "\WHFABRIC.DAT")
				ARMEDDIA1.PR_LoadFabricFromFile(g_sPathJOBST & "\TEMPLTS\POWERNET.DAT")
			End If
			
			If g_JOBSTEX = True Or g_JOBSTEX_FL = True Then
				'Jobstex fabric
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
			End If
			
			'If we have values for Ankle Pressure and fabric lets use them
			If g_iLtLastFabric < cboFabric.Items.Count Then
				cboFabric.SelectedIndex = g_iLtLastFabric
				g_iLtLastFabric = -1000 'This will force re-figuring
				If g_iLtMM(g_iLtAnkle) > 0 Then
					txtLeftMM(g_iLtAnkle).Text = CStr(g_iLtMM(g_iLtAnkle))
					PR_FigureLeftAnkle()
				End If
				g_iLtLastFabric = -1
			End If
			
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optType.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optType_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optType.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optType.GetIndex(eventSender)
			
			Dim ii As Short
			Dim nValue As Double
			
			g_iLastLegStyle = g_iLegStyle
			g_iLegStyle = Index
			
			'get the Style String g_sStyleString from the relevent DDE box
			'or rebuild string from existing displayed data
			Select Case g_iLegStyle
				Case 0
					g_sStyleString = txtAnklet.Text
					'            g_sBestCatNo = "0105"   'Anklet
					'            g_nOptions = 0
				Case 1
					g_sStyleString = txtKneeLength.Text
					'            g_sBestCatNo = "0101"   'Knee length
					'            g_nOptions = 0
				Case 2
					g_sStyleString = txtThighLength.Text
					'            g_sBestCatNo = "0201"   'Thigh length
					'            g_nOptions = 0
				Case 3
					g_sStyleString = txtKneeBand.Text
					'            g_sBestCatNo = "0015"   'Knee band
					'            g_nOptions = 3
					'            g_sCatNo(1) = "0015"    'Knee band
					'            g_sCatNo(2) = "1131"    'Stump Support Below Knee
					'            g_sCatNo(3) = "0101"    'Footless Knee Length
				Case 4
					g_sStyleString = txtThighBandAK.Text
					'            g_sBestCatNo = "0019"   'Thigh band
					'            g_nOptions = 3
					'            g_sCatNo(1) = "0019"    'Thigh band
					'            g_sCatNo(2) = "1130"    'Stump Support above Knee
					'            g_sCatNo(3) = "0201"    'Footless thigh Length
					
				Case 5
					g_sStyleString = txtThighBandBK.Text
					'            g_sBestCatNo = "0019"    'Thigh band
					'            g_nOptions = 3
					'            g_sCatNo(1) = "0019"    'Thigh band
					'            g_sCatNo(2) = "1131"    'Stump Support Below Knee
					'            g_sCatNo(3) = "0201"    'Footless thigh Length
					
			End Select
			
			'Find First and last Tapes as displayed
			'Don't worry about holes for now
			g_iFirstTape = -1
			For ii = 0 To 29
				If Val(txtLeft(ii).Text) > 0 Then Exit For
			Next ii
			If ii < 30 Then g_iFirstTape = ii
			
			g_iLastTape = -1
			For ii = 29 To 0 Step -1
				If Val(txtLeft(ii).Text) > 0 Then Exit For
			Next ii
			If ii >= 0 Then g_iLastTape = ii
			
			If g_iLastTape = g_iFirstTape Then
				g_iFirstTape = -1
				g_iLastTape = -1
			End If
			
			PR_LastTapeDisplay("Enabled")
			
			Select Case Index
				Case 0 'Anklet
					'Disable FirstTape display
					PR_FirstTapeDisplay("Disabled")
					'Disable heel figuring
					'Set last tape to ankle + 1
					If g_iLtAnkle <> 0 Then
						PR_SetLastTape(g_iLtAnkle + 1)
						txtLeftMM(g_iLtAnkle).Enabled = False
						txtLeftMM(g_iLtAnkle).Text = ""
						grdLeftDisplay.Row = g_iLtAnkle - 6
						grdLeftDisplay.Col = 0
						grdLeftDisplay.Text = ""
						grdLeftDisplay.Col = 1
						grdLeftDisplay.Text = ""
						cboToeStyle.SelectedIndex = ARMEDDIA1.fnGetNumber(g_sStyleString, 12)
					End If
					
					If g_POWERNET = True Then cboLeftTemplate.SelectedIndex = 0
					
				Case 1 To 2 'Knee Length, Thigh Length
					'Disable FirstTape display
					PR_FirstTapeDisplay("Disabled")
					'Enable and figure ankle
					If g_iLtAnkle <> 0 Then
						txtLeftMM(g_iLtAnkle).Enabled = True
						cboToeStyle.SelectedIndex = ARMEDDIA1.fnGetNumber(g_sStyleString, 12)
						If g_iLtLastMM <> 0 Then
							txtLeftMM(g_iLtAnkle).Text = CStr(g_iLtLastMM)
							g_iLtLastFabric = -1000 'Use this to force a refigure
							PR_FigureLeftAnkle()
						End If
					End If
					
					'Set Last style tape from StyleString
					If g_sStyleString <> "" Then
						PR_SetLastTape(ARMEDDIA1.fnGetNumber(g_sStyleString, 3) - 1)
					End If
					
					'Ensure that the style last tape is valid, reset if not
					If g_iStyleLastTape > 0 And g_iStyleLastTape <= g_iLastTape Then
						PR_SetLastTape(g_iStyleLastTape)
					Else
						PR_SetLastTape(g_iLastTape)
					End If
					
				Case 3, 4, 5 'Knee Band, Thigh Band AK and Thigh Band BK
					'Enable FirstTape display
					PR_FirstTapeDisplay("Enabled")
					
					'Disable Heel figuring
					If g_iLtAnkle <> 0 Then
						txtLeftMM(g_iLtAnkle).Text = ""
						txtLeftMM(g_iLtAnkle).Enabled = False
						grdLeftDisplay.Row = g_iLtAnkle - 6
						grdLeftDisplay.Col = 0
						grdLeftDisplay.Text = ""
						grdLeftDisplay.Col = 1
						grdLeftDisplay.Text = ""
					End If
					
					'Set Style first tape
					'Use existing if available else use first tape as above
					'don't go past ankle
					'Set First and Last style tapes from StyleString
					If g_sStyleString <> "" Then
						PR_SetFirstTape(ARMEDDIA1.fnGetNumber(g_sStyleString, 2) - 1)
						PR_SetLastTape(ARMEDDIA1.fnGetNumber(g_sStyleString, 3) - 1)
					End If
					
					'Ensure that the style last tape is valid, reset if not
					If g_iStyleFirstTape > 0 And g_iStyleFirstTape >= g_iFirstTape Then
						PR_SetFirstTape(g_iStyleFirstTape)
					Else
						If g_iLtAnkle <> 0 Then
							PR_SetFirstTape(g_iLtAnkle)
						Else
							PR_SetFirstTape(g_iFirstTape)
						End If
					End If
					
					If g_iStyleLastTape > 0 And g_iStyleLastTape <= g_iLastTape Then
						PR_SetLastTape(g_iStyleLastTape)
					Else
						PR_SetLastTape(g_iLastTape)
					End If
					
					cboLeftTemplate.SelectedIndex = 0
					
			End Select
			
			If g_sStyleString <> "" Then
				'Restore values from style string
				chkLeftZipper.CheckState = ARMEDDIA1.fnGetNumber(g_sStyleString, 10)
				
				nValue = ARMEDDIA1.fnGetNumber(g_sStyleString, 13)
				If nValue <> 0 Then
					txtFootLength.Text = CStr(nValue)
				Else
					txtFootLength.Text = ""
				End If
				
				txtFootLength_Leave(txtFootLength, New System.EventArgs())
				
				If ARMEDDIA1.fnGetNumber(g_sStyleString, 14) <= cboLeftTemplate.Items.Count - 1 Then
					cboLeftTemplate.SelectedIndex = ARMEDDIA1.fnGetNumber(g_sStyleString, 14)
				Else
					cboLeftTemplate.SelectedIndex = -1
				End If
				If g_iLtAnkle <> 0 And g_iLegStyle > 0 And g_iLegStyle < 3 And ARMEDDIA1.fnGetNumber(g_sStyleString, 5) <> 0 Then
					txtLeftMM(g_iLtAnkle).Text = CStr(ARMEDDIA1.fnGetNumber(g_sStyleString, 5))
					g_iLtLastFabric = -1000 'Use this to force a refigure
					PR_FigureLeftAnkle()
				End If
			Else
				'Build style string from exisisting data
				g_sStyleString = FN_BuildStyleString()
			End If
			
		End If
	End Sub
	
	Private Sub PR_DisplayErrorMessage(ByRef iErrorNum As Object, ByRef sContext As Object)
		'Procedure to display error messages
		'
		'    iErrorNum   Local error number
		'    sContext    User supplied string to be displayed in the Message Label
		'
		'NOTE
		'    In many cases values returned fron functions are -ve to indicate that
		'    an error has occured.
		'    This meams that an error can be quickly established without having to
		'    know the actual error details, mearly by checking if the returned value
		'    is -ve.
		'    This works well in this particular set of modules as therv are very few
		'    time that a -ve value is required to be returned.
		'
		'    All error numbers are local to this Module.
		'
		'
		'
		
		Dim iError As Short
		Dim sError, NL As String
		
		NL = Chr(13) 'New Line character
		'Within this procedure the error number is always +ve.
		'UPGRADE_WARNING: Couldn't resolve default property of object iErrorNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		iError = System.Math.Abs(iErrorNum)
		
		sError = ""
		Select Case iError
			'Messages with repect to available Styles
			Case 500
				sError = "Anklet"
			Case 501
				sError = "Knee Length"
			Case 502
				sError = "Thigh Length"
			Case 503
				sError = "Knee Band"
			Case 504
				sError = "Thigh Band A/K"
			Case 505
				sError = "Thigh Band B/K"
				'Ankle figuring
			Case 1000
				sError = "The modulus for the chosen fabric " & VB6.GetItemString(cboFabric, cboFabric.SelectedIndex) & NL
				sError = sError & "Not listed in the POWERNET conversion chart."
			Case 1001
				sError = "The given Pressure resulted in a reduction of less than 14. "
				sError = sError & "Using a 14 reduction and back calculating the Pressure"
			Case 1002
				sError = "The given Pressure resulted in a reduction of over 26. "
				sError = sError & "As the Diagnosis is for Burns, "
				sError = sError & "using a 26 reduction and back calculating the Pressure"
			Case 1003
				sError = "The given Pressure resulted in a reduction of over 32. "
				sError = sError & "As the Diagnosis is other than Burns, "
				sError = sError & "using a 32 reduction and back calculating the Pressure"
			Case 1004
				sError = "At this reduction the 100% stretch is exceeded. "
				sError = sError & "Decrease the reduction."
			Case 2001
				sError = "Maximum Stretch exceeded. "
			Case 2002
				sError = "Minimum Stretch exceeded. "
			Case 2003
				sError = "Donning Stretch of 110% exceeded. "
			Case 2004
				sError = "Reduction exceeds 32. "
				
			Case Else
				sError = "Unknown Error" & NL
		End Select
		
		'If a message is in the box then add
		If labMessage.Text = "" Then
			labMessage.Text = sError
		Else
			labMessage.Text = labMessage.Text & NL & NL & sError
		End If
		Beep()
		
	End Sub
	
	Private Sub PR_DisplayFiguredAnkle(ByRef sLeg As String, ByRef nStretch As Double, ByRef nAnkleCir As Double, ByRef nHeelCir As Double, ByRef iFabric As Short)
		Dim nPatternDim As Double
		Dim iReduction, iMaxStretch As Short
		Dim sContext As String
		Dim iDonningStretch, iMinStretch As Short
		
		nPatternDim = (nAnkleCir / (1 + (0.01 * nStretch))) / g_nUnitsFac
		iReduction = ARMDIA1.round((1 - 1 / (0.01 * nStretch + 1)) * 100)
		iDonningStretch = ARMDIA1.round((((nHeelCir / g_nUnitsFac) - nPatternDim) / nPatternDim) * 100)
		
		grdLeftDisplay.Row = g_iLtAnkle - 6
		grdLeftDisplay.Col = 0
		'Stretch
		grdLeftDisplay.Text = Str(nStretch)
		'Reduction
		grdLeftDisplay.Col = 1
		grdLeftDisplay.Text = Str(iReduction)
		g_iLtRed(g_iLtAnkle) = iReduction
		'Max Stretch
		labLeftMaxStr.Text = CStr(iDonningStretch)
		
		
		'Set Max and Min stretches
		iMaxStretch = 44
		If iFabric <= 55 Then iMinStretch = 16 Else iMinStretch = 18
		
		If nStretch < iMinStretch Then
			PR_DisplayErrorMessage(2002, "Ankle Figuring")
		End If
		
		If nStretch > iMaxStretch Then
			PR_DisplayErrorMessage(2001, "Ankle Figuring")
		End If
		
		If iDonningStretch > 110 Then
			PR_DisplayErrorMessage(2003, "Ankle Figuring")
		End If
		
		If iReduction > 32 Then
			PR_DisplayErrorMessage(2004, "Ankle Figuring")
		End If
		
		
		
	End Sub
	
	Private Sub PR_DrawAndSaveLeg()
		'Procedure to create a Macro to draw a leg lower options
		'Anklet,  Knee High, Thigh High, Knee Band and Thigh Band
		
		Dim sFile, sBody As String
		Dim sData, sLengths, sLegStyle As String
		Dim nLegStyle, nStyleLastRed As Short
		Dim nLastTape, nFirstTape, nFabricClass As Short
		Dim nStyleLastTape, nStyleFirstTape, nAge As Short
		Dim FootLess, ii, itemplate As Short
		
		Dim nThighPltXoff, nThighPltYoff As Double
		Dim nThighPltRad, nTopThigh As Double
		Dim nThighPltDeltaAngle As Double
		Dim nHeel, nThighPltLen, nThighTopExtension As Double
		Dim nB, nA, nThighPltStartAngle As Double
		
		Dim nFoldHt, nValue As Double
		
		'Open Macro file (fNum is declared as Global)
		sFile = "C:\JOBST\DRAW.D"
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fnum = FN_DrawOpen(sFile, txtPatientName, txtFileNo, txtLeg)
		
		PR_SaveLeg()
		
		BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\LEG\LG_LEG.D;")
		
		
		'Patient Details
		BDYUTILS.PR_PutStringAssign("sPatient", txtPatientName)
		BDYUTILS.PR_PutStringAssign("sFileNo", txtFileNo)
		
		If txtWorkOrder.Text = "" Then
			BDYUTILS.PR_PutStringAssign("sWorkOrder", "-")
		Else
			BDYUTILS.PR_PutStringAssign("sWorkOrder", txtWorkOrder)
		End If
		
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
		
		'Body details
		BDYUTILS.PR_PutStringAssign("sFabric", txtFabric)
		
		'Leg Details
		sLengths = txtLeftLengths.Text 'Use later to establish 1st & Last tapes
		
		If Len(txtFabric.Text) > 0 Then BDYUTILS.PR_PutStringAssign("sFabric", txtFabric)
		
		If ARMEDDIA1.fnGetNumber(g_sStyleString, 4) < 0 Or ARMEDDIA1.fnGetNumber(g_sStyleString, 1) > 2 Or g_iLtAnkle = 0 Then
			BDYUTILS.PR_PutLine("FootLess = %true;")
			FootLess = True
		Else
			BDYUTILS.PR_PutLine("FootLess = %false;")
			'Ankle tape values
			BDYUTILS.PR_PutNumberAssign("nAnkleTape", ARMEDDIA1.fnGetNumber(g_sStyleString, 4))
			BDYUTILS.PR_PutStringAssign("sAnkleTape", Str(ARMEDDIA1.fnGetNumber(g_sStyleString, 4)))
			BDYUTILS.PR_PutStringAssign("sMMAnkle", Str(ARMEDDIA1.fnGetNumber(g_sStyleString, 5)))
			BDYUTILS.PR_PutStringAssign("sGramsAnkle", Str(ARMEDDIA1.fnGetNumber(g_sStyleString, 6)))
			BDYUTILS.PR_PutStringAssign("sReductionAnkle", Str(ARMEDDIA1.fnGetNumber(g_sStyleString, 7)))
			BDYUTILS.PR_PutNumberAssign("nReductionAnkle", ARMEDDIA1.fnGetNumber(g_sStyleString, 7))
			nHeel = ARMEDDIA1.fnGetNumber(g_sStyleString, 9)
		End If
		
		nFabricClass = ARMEDDIA1.fnGetNumber(g_sStyleString, 11)
		If nFabricClass = 2 Then
			BDYUTILS.PR_PutStringAssign("sReduction", txtLeftRed)
			BDYUTILS.PR_PutStringAssign("sTapeMMs", txtLeftMMs)
			BDYUTILS.PR_PutStringAssign("sStretch", txtLeftStr)
		End If
		
		BDYUTILS.PR_PutStringAssign("sLeg", txtLeg)
		
		BDYUTILS.PR_PutStringAssign("sPressure", txtLeftTemplate)
		itemplate = CShort(VB.Left(txtLeftTemplate.Text, 2))
		
		BDYUTILS.PR_PutStringAssign("sTapeLengths", txtLeftLengths)
		BDYUTILS.PR_PutStringAssign("sToeStyle", txtToeStyle)
		
		BDYUTILS.PR_PutNumberAssign("nFootPleat1", ARMEDDIA1.fnDisplaytoInches(Val(txtFootPleat1.Text)))
		BDYUTILS.PR_PutNumberAssign("nFootPleat2", ARMEDDIA1.fnDisplaytoInches(Val(txtFootPleat2.Text)))
		
		BDYUTILS.PR_PutStringAssign("sFootLength", txtFootLength)
		
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
		
		'First and last tapes for style
		nStyleFirstTape = ARMEDDIA1.fnGetNumber(g_sStyleString, 2)
		BDYUTILS.PR_PutNumberAssign("nStyleFirstTape", nStyleFirstTape)
		
		nStyleLastTape = ARMEDDIA1.fnGetNumber(g_sStyleString, 3)
		BDYUTILS.PR_PutNumberAssign("nStyleLastTape", nStyleLastTape)
		
		'Get last tape value for use in calculating thigh length ending
		nValue = ARMEDDIA1.fnDisplaytoInches(Val(Mid(sLengths, ((nStyleLastTape - 1) * 4) + 1, 4)) / 10)
		
		'Style type etc
		'NOTE for a footless thigh length treat as a thigh band (BK)
		nLegStyle = ARMEDDIA1.fnGetNumber(g_sStyleString, 1)
		If nLegStyle = 2 And FootLess Then nLegStyle = 5
		
		BDYUTILS.PR_PutNumberAssign("nLegStyle", nLegStyle)
		BDYUTILS.PR_PutStringAssign("sType", g_sStyleString)
		BDYUTILS.PR_PutNumberAssign("nTopLegPleat1", ARMEDDIA1.fnDisplaytoInches(Val(txtTopLegPleat1.Text)))
		BDYUTILS.PR_PutNumberAssign("nTopLegPleat2", ARMEDDIA1.fnDisplaytoInches(Val(txtTopLegPleat2.Text)))
		
		Select Case nLegStyle
			Case 1 'Knee High
				sLegStyle = "KLN"
				
				If nFabricClass = 0 And nValue < 15 Then
					nStyleLastRed = 10
				Else
					nStyleLastRed = 14
				End If
				
			Case 2 'Thigh Length
				sLegStyle = "TLN"
				
				If nFabricClass = 0 Then
					nStyleLastRed = 14
				Else
					nStyleLastRed = 16
				End If
				
				'Calculate parameters for Thigh Plate Template
				nTopThigh = ((nValue * ((100 - nStyleLastRed) / 100)) / 2) + 0.1875
				nThighPltRad = 23
				
				Select Case nTopThigh 'Get closest thigh plate
					Case 0 To 7
						nThighPltXoff = 0.52
						nThighPltLen = 6
					Case 7 To 9
						nThighPltXoff = 0.73
						nThighPltLen = 8
					Case 9 To 11
						nThighPltXoff = 1
						nThighPltLen = 10
					Case 11 To 13
						nThighPltXoff = 1.27
						nThighPltLen = 12
					Case Is > 13
						nThighPltXoff = 1.5
						nThighPltLen = 14
				End Select
				
				'Depending on Heel set top of thigh extension
				If nHeel < 9 Then
					nThighTopExtension = 0.5
					BDYUTILS.PR_PutNumberAssign("nThighTopExtension", nThighTopExtension)
				Else
					nThighTopExtension = 1
					BDYUTILS.PR_PutNumberAssign("nThighTopExtension", nThighTopExtension)
				End If
				
				'Revise thigh plate to make it fit (if required)
				'Note the thigh plate line must land on the profile, on or
				'in front of the last tape
				If nThighPltXoff < nThighTopExtension Then nThighPltXoff = 1
				
				BDYUTILS.PR_PutNumberAssign("nThighPltRad", nThighPltRad)
				BDYUTILS.PR_PutNumberAssign("nThighPltXoff", nThighPltXoff)
				
			Case 3 'Knee Band
				sLegStyle = "KBN"
				
				BDYUTILS.PR_PutNumberAssign("nStyleFirstRed", 8)
				nStyleLastRed = 8
				
				'Elastic at distal always
				BDYUTILS.PR_PutNumberAssign("nElastic", 1)
				
			Case 4, 5 'Thigh Band Above Knee and ThighBand Below Knee
				If nFabricClass = 0 Then
					nStyleLastRed = 14
				Else
					nStyleLastRed = 16
				End If
				
				'Release the distal tape to a 95% reduction,  92% reduction if +3 or +1-1/2 tape
				If nStyleFirstTape > 8 Then
					BDYUTILS.PR_PutNumberAssign("nStyleFirstRed", 5)
				Else
					BDYUTILS.PR_PutNumberAssign("nStyleFirstRed", 8)
				End If
				
				'Calculate parameters for Thigh Plate Template
				nTopThigh = ((nValue * ((100 - nStyleLastRed) / 100)) / 2) + 0.1875
				nThighPltRad = 23
				
				Select Case nTopThigh 'Get closest thigh plate
					Case 0 To 7
						nThighPltXoff = 0.52
						nThighPltLen = 6
					Case 7 To 9
						nThighPltXoff = 0.73
						nThighPltLen = 8
					Case 9 To 11
						nThighPltXoff = 1
						nThighPltLen = 10
					Case 11 To 13
						nThighPltXoff = 1.27
						nThighPltLen = 12
					Case Is > 13
						nThighPltXoff = 1.5
						nThighPltLen = 14
				End Select
				
				'Depending on age set top of thigh extension
				If nAge < 10 Then
					nThighTopExtension = 0.5
					BDYUTILS.PR_PutNumberAssign("nThighTopExtension", nThighTopExtension)
				Else
					nThighTopExtension = 1
					BDYUTILS.PR_PutNumberAssign("nThighTopExtension", nThighTopExtension)
				End If
				
				'Revise thigh plate to make it fit (if required)
				'Note the thigh plate line must land on the profile, on or
				'in front of the last tape
				If nThighPltXoff < nThighTopExtension Then nThighPltXoff = 1
				
				BDYUTILS.PR_PutNumberAssign("nThighPltRad", nThighPltRad)
				BDYUTILS.PR_PutNumberAssign("nThighPltXoff", nThighPltXoff)
				
				'Elastic
				If nLegStyle = 4 Then
					BDYUTILS.PR_PutNumberAssign("nElastic", 1)
					sLegStyle = "TBA" 'Thigh Band (A/K)
				Else
					BDYUTILS.PR_PutNumberAssign("nElastic", 0)
					sLegStyle = "TBB" 'Thigh Band (B/K)
				End If
				
				
			Case Else 'Anklet,
				sLegStyle = "ANK"
				'Anklet reductions done by counting and therefor done in the
				'macro LGLEGDWG.D
				nStyleLastRed = -1
				BDYUTILS.PR_PutNumberAssign("nReductionAnkle", 14)
				
		End Select
		
		BDYUTILS.PR_PutNumberAssign("nStyleLastRed", nStyleLastRed)
		BDYUTILS.PR_PutStringAssign("sLegStyle", sLegStyle)
		
		'Fabric Class (Load JOBSTEX_FL Procedures and Defaults if class = 2 )
		'For a footless style the only valid classes are 0 and 1
		If FootLess And nFabricClass = 2 Then nFabricClass = 1
		BDYUTILS.PR_PutNumberAssign("nFabricClass", nFabricClass)
		
		'Load universal Procedures and Defaults
		BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\LEG\LGLEGDEF.D;")
		
		'Get Origin & Draw template
		BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\LEG\LGLEGTMP.D;")
		
		'Calculate Foot Points (If a foot exists)
		'Draw Leg and foot profile
		If Not FootLess Then
			BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\WAIST\WHFTPNTS.D;")
			BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\LEG\LGLEGDWG.D;")
		End If
		
		'Footless ThighLength and Thigh/Knee Bands
		If FootLess Then
			BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\LEG\LGBNDDWG.D;")
		End If
		
		'Draw closing lines
		Select Case nLegStyle
			Case 0
				BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\LEG\LGANKCLS.D;")
			Case 1, 2
				BDYUTILS.PR_PutLine("@" & g_sPathJOBST & "\LEG\LGLEGCLS.D;")
		End Select
		
		'Close macro file
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fnum)
		
	End Sub
	
	Private Sub PR_EnableJOBSTEX()
		'Set up the display for JOBSTEX fabric
		'
		Dim iDisable, ii, nn, iEnable As Short
		Dim nValue As Double
		
		cboLeftTemplate.Items.Clear()
		cboLeftTemplate.Items.Add("13 DS")
		cboLeftTemplate.Items.Add("09 DS")
		cboLeftTemplate.Enabled = False
		
		'Set value
		For ii = 0 To (cboLeftTemplate.Items.Count - 1)
			If VB6.GetItemString(cboLeftTemplate, ii) = txtLeftTemplate.Text Then
				cboLeftTemplate.SelectedIndex = ii
			End If
		Next ii
		
		'Set the display for leg depending on the ankle and fabric and options.
		'g_JOBSTEX_FL is a special case where all of the stretch at tapes above the
		'ankle are calculated
		'
		'Panty Legs
		'If a template is not given then
		'then set 13DS for JOBSTEX, this is first item in list
		'
		If g_iLtAnkle = 0 Then
			If cboLeftTemplate.SelectedIndex = -1 Then cboLeftTemplate.SelectedIndex = 0
			cboLeftTemplate.Enabled = False
			chkLeftZipper.Enabled = False
			chkLeftZipper.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
		
		If g_iLtAnkle <> 0 Then
			For ii = 0 To 2
				labLeftDisp(ii).Visible = True
			Next ii
			labLeftMaxStr.Visible = True
			labLeftDisp(1).Text = "Str"
			For ii = 0 To 3
				labLeftDisp(ii).Visible = True
			Next 
			'Set display value depending on Ankle position
			iEnable = g_iLtAnkle
			If iEnable = 6 Then
				iDisable = 7
			Else
				iDisable = 6
			End If
			chkLeftZipper.Enabled = True
			txtLeftMM(iEnable).Visible = True
			txtLeftMM(iDisable).Visible = False
		End If
		
		
	End Sub
	
	Private Sub PR_EnablePOWERNET()
		
		Dim iDisable, ii, nn, iEnable As Short
		Dim nValue As Double
		
		cboLeftTemplate.Items.Clear()
		cboLeftTemplate.Items.Add("30mm Hg")
		cboLeftTemplate.Items.Add("35mm Hg")
		cboLeftTemplate.Items.Add("40mm Hg")
		cboLeftTemplate.Items.Add("50mm Hg")
		
		cboLeftTemplate.Enabled = True
		
		'Set value
		For ii = 0 To (cboLeftTemplate.Items.Count - 1)
			If VB6.GetItemString(cboLeftTemplate, ii) = txtLeftTemplate.Text Then
				cboLeftTemplate.SelectedIndex = ii
			End If
		Next ii
		
		'Set the display for each leg depending on the ankle and fabric and options.
		'     g_XtAnkle = 0   => Panty leg
		'     g_XtAnkle = 6   => Ankle at +1-1/2
		'     g_XtAnkle = 0   => Ankle at +3
		'
		
		'Panty Legs
		'If a template is not given then
		'then set to 35mm Hg for Powernet
		
		If g_iLtAnkle = 0 Then
			If cboLeftTemplate.SelectedIndex = -1 Then cboLeftTemplate.SelectedIndex = 1
			chkLeftZipper.Enabled = False
			chkLeftZipper.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
		If g_iLtAnkle <> 0 Then
			For ii = 0 To 2
				labLeftDisp(ii).Visible = True
			Next ii
			labLeftDisp(1).Text = "Gms"
			labLeftDisp(3).Visible = False
			labLeftMaxStr.Visible = False
			For ii = 0 To 2
				labLeftDisp(ii).Visible = True
			Next 
			
			'Set display value depending on Ankle position
			iEnable = g_iLtAnkle
			If iEnable = 6 Then
				iDisable = 7
			Else
				iDisable = 6
			End If
			
			chkLeftZipper.Enabled = True
			
			txtLeftMM(iEnable).Visible = True
			txtLeftMM(iDisable).Visible = False
		End If
		
	End Sub
	
	Private Sub PR_EstablishAnkles()
		'For both legs
		'Setup global variables.
		
		Dim ii, nn As Short
		Dim nValue As Double
		
		'Find ankle tapes
		'Note:
		'     This depends on the heel dimension.
		'     For a heel of less than 9" the ankle is at the +1-1/2 tape
		'     otherwise it is at the +3 tape.
		'     The exception to this rule is that if one leg is at 1-1/2 and one
		'     is at +3, then both legs must be figured at the +3 tape.
		'
		'     Later a check is made to see if there are enough tapes to draw the
		'     foot properley, at this point we are not interested we just want the
		'     ankles.
		'
		'     g_XtAnkle = 0   => Panty leg
		'     g_XtAnkle = 6   => Ankle at +1-1/2
		'     g_XtAnkle = 7   => Ankle at +3
		'
		' Note also use of LAST global variables
		
		g_iLtAnkle = 0
		g_nLtLastAnkle = 0
		g_nLtLastHeel = ARMEDDIA1.fnDisplaytoInches(Val(txtLeft(5).Text))
		If g_nLtLastHeel <> 0 Then
			If g_nLtLastHeel < 9 Then
				g_iLtAnkle = 6 'Small heel Ankle at +1 1/2 tape
			Else
				g_iLtAnkle = 7 'Large heel Ankle at +3 tape
			End If
		End If
		
		g_nLtLastAnkle = ARMEDDIA1.fnDisplaytoInches(Val(txtLeft(g_iLtAnkle).Text))
		
	End Sub
	
	Private Sub PR_FigureLeftAnkle()
		'Figures the Pressure/Stretch at the LEFT ankle
		'Input from form:-
		'   Left(5-7) where     Left(5) = Heel Tape
		'                       Left(6) = Ankle Tape (Heel < 9")
		'                       Left(7) = Ankle Tape (Heel > 9")
		'   cboFabric
		'   chkLeftZipper
		'   txtLeftMM(g_iLtAnkle)
		'   txtLeftStretch(g_iLtAnkle)   JOBSTEX fabric only
		'
		'Input from globals
		'   g_JOBSTEX           Flag to indicate use of JOBSTEX fabric
		'   g_JOBSTEX_FL        Flag to indicate that JOBSTEX fabric in use
		'                       and that the stretch will be calculated at all
		'                       leg tape
		'   g_POWERNET          Flag to indicate use of Powernet fabric
		'
		'   g_iLtAnkle          Indicates the Left ankle tape
		'   g_nLtLastHeel
		'   g_nLtLastAnkle
		'   g_iLtLastMM
		'   g_iLtLastStretch    JOBSTEX fabric only
		'   g_iLtLastZipper
		'   g_iLastFabric
		'
		'JOBSTEX fabric Input and Output
		'   txtLeftMM(g_iLtAnkle)        User Input, if txtLeftStretch has been modified.
		'   txtLeftStretch(g_iLtAnkle)   User Input, if txtLeftMM has been modified.
		'   labLeftRed(g_iLtAnkle)       Displays Reduction
		'   labLeftMaxStr(2)             Displays MaxStretch
		'
		'Note:-
		'   The user inputs a required pressure in MMHg the stretch is then calculated
		'   and the stretch dependant values displayed.
		'   Alternately the stretch can be entered and the pressure in MMHg calculated.
		'   Given input of both (eg at link close) then MMHg takes precidence
		'
		'POWERNET fabric Input and Output
		'   txtLeftMM(g_iLtAnkle)            User Input
		'   txtLeftStretch(g_iLtAnkle)       not used (disabled)
		'   labLeftFiguredGrams(g_iLtAnkle)  Displays Grams
		'   labLeftRed(g_iLtAnkle)           Displays Reduction
		'Note:-
		'   The user inputs a required pressure in MMHg the reduction is then calculated
		'   and the reduction dependant values displayed.
		'
		'
		Dim nMMHg, nStretch As Double
		Dim nAnkleCir, nHeelCir As Double
		Dim nTapeCir As Double
		Dim iFabric, iZipper, iPrevMMHg As Short
		Dim iGrams, iModulus, iReduction As Short
		Dim ii As Short
		Dim nLastTape, nLastTapeMMHg As Short
		Dim nLastTapeCir As Double
		
		'Clean Messaage box
		labMessage.Text = ""
		
		'Establish if enough data exists to calculate.
		'If not then EXIT quietly
		
		'Leg tapes at heel and ankles and Fabric
		If Val(txtLeft(5).Text) = 0 Or Val(txtLeft(g_iLtAnkle).Text) = 0 Or cboFabric.SelectedIndex = -1 Then
			Exit Sub
		End If
		
		'MMHg for POWERNET fabric
		If g_POWERNET = True And Val(txtLeftMM(g_iLtAnkle).Text) = 0 Then Exit Sub
		
		'MMHg and Stretch for JOBSTEX fabric
		If (g_JOBSTEX = True Or g_JOBSTEX_FL = True) And Val(txtLeftMM(g_iLtAnkle).Text) = 0 Then Exit Sub
		
		'Fabric Selection
		'NOTE:-
		'    The POWERNET fabric is based on the file g_sPathJOBST & "\WHFABRIC.DAT
		'    For clarity this file can contain blank lines.  Therefor we must check
		'    that the fabric selected is not blank.
		If VB6.GetItemString(cboFabric, cboFabric.SelectedIndex) = "" Then Exit Sub
		
		'Establish if any of the inputs have changed compared to the last stored
		'values.
		'If they have changed then recalculate and redisplay
		'NB  Use of GOTO, (structured programmers get stuffed!)
		
		'Heel Tape
		If ARMEDDIA1.fnDisplaytoInches(Val(txtLeft(5).Text)) <> g_nLtLastHeel Then GoTo CALC
		
		'Ankle Tape
		If ARMEDDIA1.fnDisplaytoInches(Val(txtLeft(g_iLtAnkle).Text)) <> g_nLtLastAnkle Then GoTo CALC
		
		'Pressure MMHg
		If Val(txtLeftMM(g_iLtAnkle).Text) <> g_iLtLastMM Then GoTo CALC
		
		'Zippers
		If chkLeftZipper.CheckState <> g_iLtLastZipper Then GoTo CALC
		
		'Fabric
		If cboFabric.SelectedIndex <> g_iLtLastFabric Then GoTo CALC
		
		'If nothing has changed then Exit this sub
		Exit Sub
		
		'If we get to here we can assume that there is enough data to process and
		'we can calculate the relevant values.
CALC: 
		'Calculate stretch based on given pressure (MMHg).
		'NB  All dimensions are converted to in inches except that last dimesions
		'    are stored using their display values.
		
		nHeelCir = ARMEDDIA1.fnDisplaytoInches(Val(txtLeft(5).Text))
		g_nLtLastHeel = nHeelCir
		
		nAnkleCir = ARMEDDIA1.fnDisplaytoInches(Val(txtLeft(g_iLtAnkle).Text))
		g_nLtLastAnkle = nAnkleCir
		
		iZipper = chkLeftZipper.CheckState
		g_iLtLastZipper = iZipper
		
		g_iLtLastFabric = cboFabric.SelectedIndex
		nLastTape = -1
		
		nMMHg = Val(txtLeftMM(g_iLtAnkle).Text)
		If g_JOBSTEX = True Or g_JOBSTEX_FL = True Then
			'JOBSTEX - Fabric
			iFabric = Val(VB.Left(VB6.GetItemString(cboFabric, cboFabric.SelectedIndex), 2))
			nStretch = ARMDIA1.round(FN_JOBSTEX_Stretch(nAnkleCir, nHeelCir, nMMHg, iZipper, iFabric))
			
			'Calculate and Display the other stretch dependent values
			PR_DisplayFiguredAnkle("Left", nStretch, nAnkleCir, nHeelCir, iFabric)
			
			g_iLtLastStretch = nStretch
			
			'g_iLtRed(g_iLtAnkle) this is set in PR_DisplayFiguredAnkle above
			g_iLtStretch(g_iLtAnkle) = nStretch
			g_iLtMM(g_iLtAnkle) = nMMHg
			
			
			'Check if template is set
			If cboLeftTemplate.SelectedIndex = -1 Then cboLeftTemplate.SelectedIndex = iZipper
			
		Else
			'POWERNET - Fabric
			iGrams = ARMDIA1.round(nAnkleCir * nMMHg)
			iModulus = Val(Mid(VB6.GetItemString(cboFabric, cboFabric.SelectedIndex), 5, 3))
			iReduction = FN_POWERNET_Reduction(iGrams, iModulus)
			'Exit if error
			If iReduction < 0 Then
				PR_DisplayErrorMessage(iReduction, "Error - Left Ankle Figuring")
				Exit Sub
			End If
			
			'Minimum Reduction for all diagnosis
			If iReduction < 14 Then
				iReduction = 14
				nMMHg = FN_POWERNET_Pressure(iReduction, nAnkleCir, iModulus)
				iGrams = ARMDIA1.round(nAnkleCir * nMMHg)
				PR_DisplayErrorMessage(1001, "Left Ankle Figuring")
			End If
			
			'Maximum Reduction for Burns
			If iReduction > 26 And txtDiagnosis.Text = "Burns" Then
				iReduction = 26
				nMMHg = FN_POWERNET_Pressure(iReduction, nAnkleCir, iModulus)
				iGrams = ARMDIA1.round(nAnkleCir * nMMHg)
				PR_DisplayErrorMessage(1002, "Left Ankle Figuring")
			End If
			
			'Maximum reduction for all other diagnosis
			If iReduction > 32 Then
				iReduction = 32
				nMMHg = FN_POWERNET_Pressure(iReduction, nAnkleCir, iModulus)
				iGrams = ARMDIA1.round(nAnkleCir * nMMHg)
				PR_DisplayErrorMessage(1003, "Left Ankle Figuring")
			End If
			
			'Display Calculations
			txtLeftMM(g_iLtAnkle).Text = Str(nMMHg)
			grdLeftDisplay.Row = g_iLtAnkle - 6
			grdLeftDisplay.Col = 0
			grdLeftDisplay.Text = Str(iGrams)
			grdLeftDisplay.Col = 1
			grdLeftDisplay.Text = Str(iReduction)
			
			'Set Template
			Select Case iReduction
				Case 0 To 14, 14 To 18
					cboLeftTemplate.SelectedIndex = 0 '30 mmhg
				Case 19 To 21
					cboLeftTemplate.SelectedIndex = 1 '35 mmhg
				Case 21 To 23
					cboLeftTemplate.SelectedIndex = 2 '40 mmhg
				Case 24 To 32, Is > 32
					cboLeftTemplate.SelectedIndex = 3 '50 mmhg
			End Select
			
			'Max Stretch (Only if no zipper given)
			If chkLeftZipper.CheckState = 0 Then
				nHeelCir = nHeelCir / 4
				nAnkleCir = (nAnkleCir * ((100 - iReduction) / 100)) / 2
				If nAnkleCir < nHeelCir Then PR_DisplayErrorMessage(1004, "Left Ankle Figuring")
			End If
			
			g_iLtStretch(g_iLtAnkle) = iGrams 'Bit of a misnomer but what the Hell!
			g_iLtRed(g_iLtAnkle) = iReduction
			g_iLtMM(g_iLtAnkle) = nMMHg
			
		End If
		g_iLtLastMM = Val(txtLeftMM(g_iLtAnkle).Text)
		
		'Calculate the pressure at all leg tapes
		'
		If g_JOBSTEX_FL = True Then
			'iZipper = 0     'Don't use zipper
			
			'Establish Last tape position
			'This allows for the addition of leg tapes to extend up the leg
			'It's not very efficient but what the hell!!!!
			nLastTape = 0
			For ii = g_iLtAnkle + 1 To 29
				If txtLeft(ii).Text = "" Then Exit For
			Next ii
			If ii = 29 Then
				nLastTape = 29
			Else
				nLastTape = ii - 1
			End If
			
			'Establish pressure at last tape
			'As the last tape is at a 20 reduction we know that the stretch must
			'be 25%
			nLastTapeCir = ARMEDDIA1.fnDisplaytoInches(Val(txtLeft(nLastTape).Text))
			nLastTapeMMHg = ARMDIA1.round(FN_JOBSTEX_Pressure(nLastTapeCir, nHeelCir, 25, iZipper, iFabric))
			
			txtLeftMM(nLastTape).Text = CStr(nLastTapeMMHg)
			g_iLtStretch(nLastTape) = 25
			g_iLtRed(nLastTape) = 20
			g_iLtMM(nLastTape) = nLastTapeMMHg
			
			'First pass from Ankle to LastTape
			nStretch = g_iLtLastStretch
			iPrevMMHg = g_iLtLastMM
			For ii = g_iLtAnkle + 1 To nLastTape - 1
				nTapeCir = ARMEDDIA1.fnDisplaytoInches(Val(txtLeft(ii).Text))
				nMMHg = 1000 'Force While
				While nMMHg > iPrevMMHg And nStretch > 1
					nMMHg = ARMDIA1.round(FN_JOBSTEX_Pressure(nTapeCir, nHeelCir, nStretch, iZipper, iFabric))
					If nMMHg > iPrevMMHg Then nStretch = nStretch - 1
				End While
				txtLeftMM(ii).Text = CStr(nMMHg)
				g_iLtMM(ii) = nMMHg
				iPrevMMHg = nMMHg
				g_iLtStretch(ii) = nStretch
				g_iLtRed(ii) = ARMDIA1.round((1 - 1 / (0.01 * nStretch + 1)) * 100)
			Next ii
			
			'Second pass from LastTape to Ankle
			'Based on last tape Pressure recalculate the tapes until the pressures
			'become the same
			iPrevMMHg = nLastTapeMMHg
			
			For ii = nLastTape - 1 To g_iLtAnkle + 1 Step -1
				nTapeCir = ARMEDDIA1.fnDisplaytoInches(Val(txtLeft(ii).Text))
				
				nMMHg = Val(txtLeftMM(ii).Text)
				nStretch = g_iLtStretch(ii)
				If nMMHg >= iPrevMMHg Then Exit For
				While nMMHg < iPrevMMHg
					nMMHg = ARMDIA1.round(FN_JOBSTEX_Pressure(nTapeCir, nHeelCir, nStretch, iZipper, iFabric))
					If nMMHg < iPrevMMHg Then nStretch = nStretch + 1
				End While
				txtLeftMM(ii).Text = CStr(nMMHg)
				g_iLtMM(ii) = nMMHg
				iPrevMMHg = nMMHg
				g_iLtStretch(ii) = nStretch
				g_iLtRed(ii) = ARMDIA1.round((1 - 1 / (0.01 * nStretch + 1)) * 100)
			Next ii
			
			'Display Values
			'STRETCH
			grdLeftDisplay.Col = 0
			For ii = g_iLtAnkle + 1 To nLastTape
				grdLeftDisplay.Row = ii - 6
				grdLeftDisplay.Text = Str(g_iLtStretch(ii))
			Next ii
			
			'REDUCTION
			grdLeftDisplay.Col = 1
			For ii = g_iLtAnkle + 1 To nLastTape
				grdLeftDisplay.Row = ii - 6
				grdLeftDisplay.Text = Str(g_iLtRed(ii))
			Next ii
			
		End If
		
	End Sub
	
	Private Sub PR_FigureLeftTape(ByRef Index As Short)
		'Figures the Pressure/Stretch at a LEFT LegTape
		'This is only available to JOBSTEX_FL
		'See PR_FigureLeftTape for details
		
		Dim nMMHg, nStretch As Double
		Dim nTapeCir As Double
		Dim iZipper, iFabric As Short
		
		'Establish if enough data exists to calculate.
		'If not then EXIT quietly
		If Val(txtLeftMM(Index).Text) <= 0 Then Exit Sub
		
		'Establish if pressure has changed from last time
		'If it hasn't then exit
		'This is used to cure a problem in that the stretch for a tape is not based not on the
		'pressure but on the previous or subsequent stretch, depending if it was found
		'on the forward pass or the backward pass.
		'The pressure was then back calculated fron the stretch.
		'We had a problem in that the same pressue can give a different stretch
		'To stop this happening when tabbing through the pressures this check has
		'been introduced
		'
		If Val(txtLeftMM(Index).Text) = g_iLtMM(Index) Then Exit Sub
		
		
		'MMHg and Stretch for JOBSTEX fabric
		'Don't Calculate if no values for Ankle tape
		If (g_JOBSTEX = True Or g_JOBSTEX_FL = True) And Val(txtLeftMM(g_iLtAnkle).Text) = 0 Or g_iLtStretch(g_iLtAnkle) = 0 Then Exit Sub
		
		
		'If we get to here we can assume that there is enough data to process and
		'we can calculate the relevant values.
		
		nTapeCir = ARMEDDIA1.fnDisplaytoInches(Val(txtLeft(Index).Text))
		iZipper = chkLeftZipper.CheckState
		iFabric = Val(VB.Left(VB6.GetItemString(cboFabric, cboFabric.SelectedIndex), 2))
		nMMHg = Val(txtLeftMM(Index).Text)
		g_iLtMM(Index) = nMMHg
		
		'Calculate stretch
		nStretch = FN_JOBSTEX_Stretch(nTapeCir, g_nLtLastHeel, nMMHg, iZipper, iFabric)
		
		'Store values
		g_iLtStretch(Index) = ARMDIA1.round(nStretch)
		g_iLtRed(Index) = ARMDIA1.round((1 - 1 / (0.01 * nStretch + 1)) * 100)
		
		'Display values
		'Note, the display grid displays from the ankle only
		'
		grdLeftDisplay.Row = Index - 6 'Display from ankle
		grdLeftDisplay.Col = 0
		grdLeftDisplay.Text = Str(g_iLtStretch(Index))
		grdLeftDisplay.Col = 1
		grdLeftDisplay.Text = Str(g_iLtRed(Index))
		
	End Sub
	
	Private Sub PR_FirstTapeDisplay(ByRef sType As String)
		
		If sType = "Disabled" Then
			txtFirstTape.Text = ""
			txtFirstTape.Enabled = False
			labFirstTape.Enabled = False
			If g_iStyleFirstTape >= 0 And g_iStyleFirstTape <= 29 Then Label1(g_iStyleFirstTape).ForeColor = System.Drawing.ColorTranslator.FromOle(RGB(0, 0, 0))
			spnFirstTape.Enabled = False
		Else
			txtFirstTape.Enabled = True
			labFirstTape.Enabled = True
			spnFirstTape.Enabled = True
		End If
		
	End Sub
	
	Private Sub PR_HeelChange(ByRef sLeg As String)
		'Procedure to modify the display if heel changed
		
		Dim iLtExistingAnkle, iRtExistingAnkle As Short
		'Clean Message box
		labMessage.Text = ""
		
		'Exit if heel has not been changed
		If ARMEDDIA1.fnDisplaytoInches(Val(txtLeft(5).Text)) = g_nLtLastHeel Then Exit Sub
		
		'Store existing ankle position and get new ones
		iLtExistingAnkle = g_iLtAnkle
		
		PR_EstablishAnkles()
		
		'If revised the same as existing then exit
		If iLtExistingAnkle = g_iLtAnkle Then Exit Sub
		
		'Modify display
		PR_ResetFormLGLEGDIA()
		
		'Disable / Enable Anklet | Knee
		If g_iLtAnkle = 0 Then
			optType(0).Enabled = False
			optType(1).Enabled = False
			optType(2).Enabled = False
		Else
			optType(0).Enabled = True
			optType(1).Enabled = True
			optType(2).Enabled = True
		End If
		
		If g_POWERNET = True Then PR_EnablePOWERNET()
		If g_JOBSTEX = True Or g_JOBSTEX_FL = True Then PR_EnableJOBSTEX()
		' If g_JOBSTEX_FL = True Then PR_EnableFL_JOBSTEX
		
	End Sub
	
	Private Sub PR_LastTapeDisplay(ByRef sType As String)
		If sType = "Disabled" Then
			txtLastTape.Text = ""
			txtLastTape.Enabled = False
			labLastTape.Enabled = False
			If g_iStyleLastTape >= 0 And g_iStyleLastTape <= 29 Then Label1(g_iStyleLastTape).ForeColor = System.Drawing.ColorTranslator.FromOle(RGB(0, 0, 0))
			spnLastTape.Enabled = False
		Else
			txtLastTape.Enabled = True
			labLastTape.Enabled = True
			spnLastTape.Enabled = True
		End If
		
	End Sub
	
	Private Sub PR_LoadFabricFromFile(ByRef sFileName As String)
		'Procedure to load the POWERNET conversion chart from file
		'N.B. File opening Errors etc. are not handled (so tough titty!)
		
		Dim fnum, ii As Short
		fnum = FreeFile
		FileOpen(fnum, sFileName, OpenMode.Input)
		ii = 0
		Do Until EOF(fnum)
			Input(fnum, POWERNET.Modulus(ii))
			Input(fnum, POWERNET.Conversion_Renamed(ii))
			ii = ii + 1
		Loop 
		FileClose(fnum)
		
	End Sub
	
	Private Sub PR_PutLine(ByRef sLine As String)
		'Puts the contents of sLine to the opened "Macro" file
		'Puts the line with no translation or additions
		'    fNum is global variable
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, sLine)
	End Sub
	
	Private Sub PR_PutNumberAssign(ByRef sVariableName As String, ByRef nAssignedNumber As Object)
		
		'Procedure to put a number assignment
		'Adds a semi-colon
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nAssignedNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, sVariableName & "=" & Str(nAssignedNumber) & ";")
		
		
	End Sub
	
	Private Sub PR_PutStringAssign(ByRef sVariableName As String, ByRef sAssignedString As Object)
		'Procedure to put a string assignment
		'Encloses String in quotes and adds a semi-colon
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, sVariableName & "=" & QQ & sAssignedString & QQ & ";")
		
	End Sub
	
	Private Sub PR_ResetFormLGLEGDIA()
		'Reset display of form to a clean state
		Dim ii As Short
		
		'Reset Max Stretch
		labLeftMaxStr.Text = ""
		labLeftMaxStr.Visible = False
		labLeftDisp(3).Visible = False
		
		'Clean Display grid
		For ii = 0 To 1 'First two rows only, used in this case
			
			grdLeftDisplay.Row = ii
			grdLeftDisplay.Col = 0
			grdLeftDisplay.Text = ""
			grdLeftDisplay.Col = 1
			grdLeftDisplay.Text = ""
			
		Next ii
		
		'Switch off MM text boxes
		For ii = 6 To 7
			txtLeftMM(ii).Text = ""
			txtLeftMM(ii).Visible = False
		Next ii
		
		'
		For ii = 0 To 2
			labLeftDisp(ii).Visible = False
		Next ii
		
	End Sub
	
	Private Sub PR_SaveLeg()
		'Procedure to create a macro to save the leg data
		If txtUidLeftLeg.Text <> "" Then
			BDYUTILS.PR_PutLine("HANDLE  hLeg;")
			'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fnum, "hLeg = UID (" & QQ & "find" & QC & Val(txtUidLeftLeg.Text) & ");")
		Else
			BDYUTILS.PR_PutLine("HANDLE  hLeg, hTitle, hLegTitle;")
			BDYUTILS.PR_PutLine("XY xyTitleOrigin, xyTitleScale;")
			BDYUTILS.PR_PutLine("ANGLE aTitleAngle;")
			'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fnum, "hTitle = UID (" & QQ & "find" & QC & Val(txtUidTitle.Text) & ");")
			BDYUTILS.PR_PutStringAssign("sLeg", txtLeg)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fnum, "SetSymbolLibrary( sPathJOBST + " & QQ & "\\JOBST.SLB" & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fnum, "if ( !Symbol(" & QQ & "find" & QCQ & "legleg" & QQ & "))")
			'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fnum, "Exit(%cancel, " & QQ & "Can't find >legleg< symbol to insert\nCheck your installation, that JOBST.SLB exists" & QQ & ");")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fnum, "if ( hTitle )")
			'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fnum, "GetGeometry( hTitle, &sTitleName, &xyTitleOrigin, &xyTitleScale, &aTitleAngle);")
			'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fnum, "else")
			'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fnum, "Exit(%cancel," & QQ & "Can't find > mainpatientdetails <, Use TITLE to insert Patient Data" & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fnum, "Execute(" & QQ & "menu" & QCQ & "SetLayer" & QC & "Table(" & QQ & "find" & QCQ & "layer" & QCQ & "Data" & QQ & "));")
			
			If txtUidLegTitle.Text = "" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fnum, "if ( !Symbol(" & QQ & "find" & QCQ & "legcommon" & QQ & "))")
				'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fnum, "Exit(%cancel, " & QQ & "Can't find > legcommon < symbol to insert\nCheck your installation, that JOBST.SLB exists" & QQ & ");")
				'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fnum, "hLegTitle = AddEntity(" & QQ & "symbol" & QCQ & "legcommon" & QC & "xyTitleOrigin);")
				'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fnum, "SetDBData( hLegTitle" & CQ & "Fabric" & QCQ & txtFabric.Text & QQ & ");")
				'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fnum, "SetDBData( hLegTitle" & CQ & "fileno" & QCQ & txtFileNo.Text & QQ & ");")
				'Run macro to setup DB fields
				'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fnum, "@" & g_sPathJOBST & "\LEG\LGFIELDS.D;")
				
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fnum, "if (StringCompare(" & QQ & "Left" & QC & "sLeg))")
			'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fnum, "xyTitleOrigin.x = xyTitleOrigin.x + 1.5;")
			'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fnum, "else")
			'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fnum, "xyTitleOrigin.x = xyTitleOrigin.x + 3;")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fnum, "hLeg = AddEntity(" & QQ & "symbol" & QCQ & "legleg" & QC & "xyTitleOrigin);")
			'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fnum, "SetDBData( hLeg, " & QQ & "fileno" & QCQ & txtFileNo.Text & QQ & ");")
			
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "if (!hLeg)Exit(%cancel," & QQ & "Can't find LEGBOX to Update" & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "SetDBData( hLeg, " & QQ & "Leg" & QCQ & txtLeg.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "SetDBData( hLeg, " & QQ & "TapeLengthsPt1" & QCQ & Mid(txtLeftLengths.Text, 1, 60) & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "SetDBData( hLeg, " & QQ & "TapeLengthsPt2" & QCQ & Mid(txtLeftLengths.Text, 61, 60) & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "SetDBData( hLeg, " & QQ & "Pressure" & QCQ & txtLeftTemplate.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "SetDBData( hLeg, " & QQ & "FootLength" & QCQ & txtFootLength.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "SetDBData( hLeg, " & QQ & "Fabric" & QCQ & txtFabric.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "SetDBData( hLeg, " & QQ & "FootPleat1" & QCQ & txtFootPleat1.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "SetDBData( hLeg, " & QQ & "TopLegPleat1" & QCQ & txtTopLegPleat1.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "SetDBData( hLeg, " & QQ & "FootPleat2" & QCQ & txtFootPleat2.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "SetDBData( hLeg, " & QQ & "TopLegPleat2" & QCQ & txtTopLegPleat2.Text & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "SetDBData( hLeg, " & QQ & "ToeStyle" & QCQ & txtToeStyle.Text & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "SetDBData( hLeg, " & QQ & "Anklet" & QCQ & txtAnklet.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "SetDBData( hLeg, " & QQ & "ThighLength" & QCQ & txtThighLength.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "SetDBData( hLeg, " & QQ & "KneeLength" & QCQ & txtKneeLength.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "SetDBData( hLeg, " & QQ & "ThighBand" & QCQ & txtThighBandAK.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "SetDBData( hLeg, " & QQ & "ThighBandBK" & QCQ & txtThighBandBK.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fnum, "SetDBData( hLeg, " & QQ & "KneeBand" & QCQ & txtKneeBand.Text & QQ & ");")
		
	End Sub
	
	Private Sub PR_SetFirstTape(ByRef iNewTape As Short)
		
		If iNewTape >= 0 And iNewTape <= 29 Then
			If g_iStyleFirstTape >= 0 And g_iStyleFirstTape <= 29 Then
				Label1(g_iStyleFirstTape).ForeColor = System.Drawing.ColorTranslator.FromOle(RGB(0, 0, 0))
			End If
			g_iStyleFirstTape = iNewTape
			Label1(iNewTape).ForeColor = System.Drawing.ColorTranslator.FromOle(RGB(255, 0, 0))
			txtFirstTape.Text = LTrim(Mid(g_sTextList, (iNewTape * 3) + 1, 3))
		End If
		
	End Sub
	
	Private Sub PR_SetLastTape(ByRef iNewTape As Short)
		
		If iNewTape >= 0 And iNewTape <= 29 Then
			If g_iStyleLastTape >= 0 And g_iStyleLastTape <= 29 Then
				Label1(g_iStyleLastTape).ForeColor = System.Drawing.ColorTranslator.FromOle(RGB(0, 0, 0))
			End If
			g_iStyleLastTape = iNewTape
			Label1(iNewTape).ForeColor = System.Drawing.ColorTranslator.FromOle(RGB(255, 0, 0))
			txtLastTape.Text = LTrim(Mid(g_sTextList, (iNewTape * 3) + 1, 3))
		End If
		
	End Sub
	
	Private Sub PR_SetTextData(ByRef nHoriz As Object, ByRef nVert As Object, ByRef nHt As Object, ByRef nAspect As Object, ByRef nFont As Object)
		'To the DRAFIX macro file (given by the global fNum).
		'Write the syntax to set the TEXT default attributes, these are
		'based on the values in the arguments.  Where the value is -ve then this
		'attribute is not set.
		'where :-
		'    nHoriz      Horizontal justification (1=Left, 2=Cen, 4=Right)
		'    nVert       Verticalal justification (8=Top, 16=Cen, 32=Bottom)
		'    nHt         Text height
		'    nAspect     Text aspect ratio (heigth/width)
		'    nFont       Text font (0 to 18)
		'
		'N.B. No checking is done on the values given
		'
		'Note:-
		'    fNum, CC, QQ, NL, g_nCurrTextHt, g_nCurrTextAspect,
		'    g_nCurrTextHorizJust, g_nCurrTextVertJust, g_nCurrTextFont
		'    are globals initialised by FN_Open
		'
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nHoriz. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextHorizJust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If nHoriz >= 0 And g_nCurrTextHorizJust <> nHoriz Then
			'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fnum, "SetData(" & QQ & "TextHorzJust" & QC & nHoriz & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object nHoriz. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextHorizJust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextHorizJust = nHoriz
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nVert. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextVertJust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If nVert >= 0 And g_nCurrTextVertJust <> nVert Then
			'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fnum, "SetData(" & QQ & "TextVertJust" & QC & nVert & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object nVert. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextVertJust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextVertJust = nVert
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nHt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextHt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If nHt >= 0 And g_nCurrTextHt <> nHt Then
			'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fnum, "SetData(" & QQ & "TextHeight" & QC & nHt & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object nHt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextHt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextHt = nHt
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If nAspect >= 0 And g_nCurrTextAspect <> nAspect Then
			'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fnum, "SetData(" & QQ & "TextAspect" & QC & nAspect & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object nAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextAspect = nAspect
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nFont. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextFont. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If nFont >= 0 And g_nCurrTextFont <> nFont Then
			'UPGRADE_WARNING: Couldn't resolve default property of object fnum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fnum, "SetData(" & QQ & "TextFont" & QC & nFont & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object nFont. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextFont. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextFont = nFont
		End If
		
		
	End Sub
	
	Private Sub PR_StripBadChar(ByRef Text_Box_Name As System.Windows.Forms.Control)
		'Strip the excess characters from the text boxes.
		'This is due to a DRAFIX bug with DDE Poke.
		Dim sString As String
		Dim iLength, ii As Short
		sString = Text_Box_Name.Text
		For ii = 1 To 2
			iLength = Len(sString)
			If iLength = 0 Then Exit For
			Select Case Asc(VB.Right(sString, 1))
				Case 32 To 126, 160 To 255
					Exit For
				Case Else
					sString = VB.Left(sString, iLength - 1)
			End Select
		Next ii
		Text_Box_Name.Text = sString
	End Sub
	
	Private Sub ShiftTapesDown_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ShiftTapesDown.Click
		'The effect of a single click on this button
		'is to shift all of the tapes in the direction of the
		'arrow
		Dim ii As Short
		Dim nValue As Double
		
		'Check that last tape is empty so that the previous
		'tape can be shifted into it
		If Len(txtLeft(29).Text) > 0 Then
			'Beep and exit function
			Beep()
			Exit Sub
		End If
		
		'Shift all tapes down 1
		For ii = 29 To 1 Step -1
			txtLeft(ii).Text = txtLeft(ii - 1).Text
			nValue = ARMEDDIA1.fnDisplaytoInches(Val(txtLeft(ii - 1).Text))
			grdLeftInches.Row = ii
			grdLeftInches.Text = ARMEDDIA1.fnInchesToText(nValue)
		Next ii
		
		'Initiate Heel Change (if Required)
		If Val(txtLeft(5).Text) > 0 And g_iLtAnkle = 0 Then PR_HeelChange("Left")
		If Val(txtLeft(5).Text) = 0 And g_iLtAnkle <> 0 Then PR_HeelChange("Left")
		
		'For Thigh and Knee Bands move style first and last tapes
		If g_iLegStyle = 3 Or g_iLegStyle = 4 Then
			PR_SetFirstTape(g_iStyleFirstTape + 1)
			PR_SetLastTape(g_iStyleLastTape + 1)
		End If
		
		'Clean first position
		txtLeft(0).Text = ""
		grdLeftInches.Row = 0
		grdLeftInches.Text = ""
		
	End Sub
	
	Private Sub ShiftTapesUp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ShiftTapesUp.Click
		'The effect of a single click on this button
		'is to shift all of the tapes in the direction of the
		'arrow
		Dim ii As Short
		Dim nValue As Double
		
		'Check that first tape is empty so that the following
		'tape can be shifted into it
		If Len(txtLeft(0).Text) > 0 Then
			'Beep and exit Procedure
			Beep()
			Exit Sub
		End If
		
		'Shift all tapes up 1
		For ii = 1 To 29
			txtLeft(ii - 1).Text = txtLeft(ii).Text
			nValue = ARMEDDIA1.fnDisplaytoInches(Val(txtLeft(ii).Text))
			grdLeftInches.Row = ii - 1
			grdLeftInches.Text = ARMEDDIA1.fnInchesToText(nValue)
		Next ii
		
		'Initiate Heel Change (if Required)
		If Val(txtLeft(5).Text) > 0 And g_iLtAnkle = 0 Then PR_HeelChange("Left")
		If Val(txtLeft(5).Text) = 0 And g_iLtAnkle <> 0 Then PR_HeelChange("Left")
		
		'For Thigh and Knee Bands move style first and last tapes
		If g_iLegStyle = 3 Or g_iLegStyle = 4 Then
			If g_iLtAnkle = 0 Then
				PR_SetFirstTape(g_iStyleFirstTape - 1)
			Else
				If g_iStyleFirstTape - 1 >= g_iLtAnkle Then
					PR_SetFirstTape(g_iStyleFirstTape - 1)
				Else
					PR_SetFirstTape(g_iLtAnkle)
				End If
			End If
			PR_SetLastTape(g_iStyleLastTape - 1)
		End If
		
		'Clean last position
		txtLeft(29).Text = ""
		grdLeftInches.Row = 29
		grdLeftInches.Text = ""
	End Sub
	
	Private Sub spnFirstTape_DownClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles spnFirstTape.DownClick
		
		Dim ii As Short
		If g_iStyleFirstTape <= 0 Or g_iStyleFirstTape - 1 < g_iFirstTape Then
			Beep()
			Exit Sub
		End If
		
		PR_SetFirstTape(g_iStyleFirstTape - 1)
		
	End Sub
	
	Private Sub spnFirstTape_UpClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles spnFirstTape.UpClick
		Dim ii As Short
		
		If g_iStyleFirstTape >= 28 Or g_iStyleFirstTape + 1 >= g_iStyleLastTape Then
			Beep()
			Exit Sub
		End If
		
		PR_SetFirstTape(g_iStyleFirstTape + 1)
		
	End Sub
	
	Private Sub spnLastTape_DownClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles spnLastTape.DownClick
		Dim ii As Short
		
		If g_iStyleLastTape <= 1 Or g_iStyleLastTape - 1 <= g_iStyleFirstTape Then
			Beep()
			Exit Sub
		End If
		
		PR_SetLastTape(g_iStyleLastTape - 1)
		
	End Sub
	
	Private Sub spnLastTape_UpClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles spnLastTape.UpClick
		
		Dim ii As Short
		
		If g_iStyleLastTape >= 29 Or g_iStyleLastTape + 1 > g_iLastTape Then
			Beep()
			Exit Sub
		End If
		
		PR_SetLastTape(g_iStyleLastTape + 1)
		
	End Sub
	
	Private Sub Tab_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Tab_Renamed.Click
		'Allows the user to use enter as a tab
		
		System.Windows.Forms.SendKeys.Send("{TAB}")
		
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		'Timeout timer
		'If DRAFIX to VB DDE link fails then time out
		'after 6 seconds
		End
	End Sub
	
	Private Sub txtFootLength_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFootLength.Enter
		ARMEDDIA1.Select_Text_In_Box(txtFootLength)
	End Sub
	
	Private Sub txtFootLength_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFootLength.Leave
		BDLEGDIA1.Validate_And_Display_Text_In_Box(txtFootLength, labFootLengthInInches, -1)
	End Sub
	
	Private Sub txtLeft_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLeft.Enter
		Dim Index As Short = txtLeft.GetIndex(eventSender)
		ARMEDDIA1.Select_Text_In_Box(txtLeft(Index))
	End Sub
	
	Private Sub txtLeft_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtLeft.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtLeft.GetIndex(eventSender)
		'Use enter to act as a TAB
		'  If KeyAscii = NEWLINE Then txtLeft((Index + 1) Mod 30).SetFocus
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtLeft_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLeft.Leave
		Dim Index As Short = txtLeft.GetIndex(eventSender)
		BDLEGDIA1.Validate_And_Display_Text_In_Box(txtLeft(Index), grdLeftInches, Index)
		If Index = g_iLtAnkle Then PR_FigureLeftAnkle()
		If Index = 5 Then PR_HeelChange("Left")
	End Sub
	
	Private Sub txtLeftMM_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLeftMM.Leave
		Dim Index As Short = txtLeftMM.GetIndex(eventSender)
		If Index = g_iLtAnkle Then
			PR_FigureLeftAnkle()
		Else
			PR_FigureLeftTape(Index)
		End If
	End Sub
	
	Private Sub Update_DDE_Text_Boxes()
		'Called from OK_Click and Draw
		'Update the text boxes used for DDE transfer
		Dim ii As Short
		Dim sString, sJustifiedString As String
		Dim sFabricClass As String
		Dim iStr, iRed, iMM As Short
		Dim sStr, sRed, sMM As String
		Dim nValue As Single
		
		txtFabric.Text = VB6.GetItemString(cboFabric, cboFabric.SelectedIndex)
		txtToeStyle.Text = VB6.GetItemString(cboToeStyle, cboToeStyle.SelectedIndex)
		
		'LEFT Leg
		'Assume that data has been validated earlier
		
		'Set initial values
		txtLeftLengths.Text = ""
		txtLeftRed.Text = ""
		txtLeftStr.Text = ""
		txtLeftMMs.Text = ""
		g_iLastTape = -1
		g_iFirstTape = -1
		txtLeftTemplate.Text = VB6.GetItemString(cboLeftTemplate, cboLeftTemplate.SelectedIndex)
		
		For ii = 0 To 29
			nValue = Val(txtLeft(ii).Text)
			If nValue <> 0 Then
				nValue = nValue * 10 'Shift decimal place to right
				sJustifiedString = New String(" ", 4)
				sJustifiedString = RSet(Trim(Str(nValue)), Len(sJustifiedString))
			Else
				sJustifiedString = New String(" ", 4)
			End If
			
			'Tape values
			txtLeftLengths.Text = txtLeftLengths.Text & sJustifiedString
			
			'Set first and last tape (assumes no holes in data)
			If g_iFirstTape < 0 And nValue > 0 Then g_iFirstTape = ii + 1
			If g_iLastTape < 0 And g_iFirstTape > 0 And nValue = 0 Then g_iLastTape = ii
		Next ii
		
		If g_iLastTape < 0 Then g_iLastTape = 30
		
		'Where JOBSTEX_FL has been choosen then we update the MMs, Reduction and Stretch
		'These will then be picked up by the Leg Drawing modules
		If g_JOBSTEX_FL = True And g_iLtAnkle <> 0 Then
			For ii = 0 To 29
				iMM = g_iLtMM(ii)
				iRed = g_iLtRed(ii)
				iStr = g_iLtStretch(ii)
				If iMM <> 0 Then
					sMM = New String(" ", 3)
					sMM = RSet(Trim(Str(iMM)), Len(sMM))
					sRed = New String(" ", 3)
					sRed = RSet(Trim(Str(iRed)), Len(sRed))
					sStr = New String(" ", 3)
					sStr = RSet(Trim(Str(iStr)), Len(sStr))
				Else
					sMM = New String(" ", 3)
					sRed = New String(" ", 3)
					sStr = New String(" ", 3)
				End If
				txtLeftMMs.Text = txtLeftMMs.Text & sMM
				txtLeftRed.Text = txtLeftRed.Text & sRed
				txtLeftStr.Text = txtLeftStr.Text & sStr
			Next ii
		End If
		
		g_iStyleLastTape = g_iStyleLastTape + 1
		g_iStyleFirstTape = g_iStyleFirstTape + 1
		
		g_sStyleString = FN_BuildStyleString()
		
		Select Case g_iLegStyle
			Case 0
				txtAnklet.Text = g_sStyleString
			Case 1
				txtKneeLength.Text = g_sStyleString
			Case 2
				txtThighLength.Text = g_sStyleString
			Case 3
				txtKneeBand.Text = g_sStyleString
			Case 4
				txtThighBandAK.Text = g_sStyleString
			Case 5
				txtThighBandBK.Text = g_sStyleString
		End Select
		
	End Sub
	
	Private Function Validate_Data() As Short
		'Called from OK_Click
		
		Dim NL, sError, sTextList As String
		Dim sLeftError, sRightError As String
		Dim nFirstTape, ii, nLastTape As Short
		Dim iError As Short
		Dim nLargeHeelX As Double
		Dim nFootLength, nHeelLength, nTapeX As Double
		Dim nValue As Double
		
		NL = Chr(10) 'new line
		
		'LEFT LEG Checks
		'Check Tape length data text boxes for holes (ie missing values)
		'Establish First and last tape
		
		sError = ""
		
		For ii = 0 To 29 Step 1
			If Val(txtLeft(ii).Text) > 0 Then Exit For
		Next ii
		nFirstTape = ii
		
		For ii = 29 To 0 Step -1
			If Val(txtLeft(ii).Text) > 0 Then Exit For
		Next ii
		nLastTape = ii
		
		For ii = nFirstTape To nLastTape
			If Val(txtLeft(ii).Text) = 0 Then
				sError = sError & "Missing Tape length - " & LTrim(Mid(g_sTextList, (ii * 3) + 1, 3)) & NL
			End If
		Next ii
		
		'Check that a minimum of 3 tapes are given
		If (nLastTape - nFirstTape) < 2 Then
			sError = sError & "Minimum of 3 Tapes must be given." & NL
		End If
		
		'Check if -3 tape exists for Large Heel
		If g_iLtAnkle = 7 And Val(txtLeft(3).Text) = 0 And g_iLegStyle < 3 Then
			sError = sError & "As the Heel is 9 inches and over." & NL
			sError = sError & "and there is no -3 tape the foot will not draw properly " & NL
		End If
		
		'Check if -1 1/2 tape exists for Small Heel
		If g_iLtAnkle = 6 And Val(txtLeft(4).Text) = 0 And g_iLegStyle < 3 Then
			sError = sError & "As the Heel is smaller than 9 inches." & NL
			sError = sError & "and there is no -1 1/2 tape the foot will not draw properly " & NL
		End If
		
		'Check that figuring has been done at the ankle
		'Does not apply to Anklets or Bands
		If g_iLegStyle > 0 And g_iLegStyle < 3 And g_iLtAnkle <> 0 And (g_iLtStretch(g_iLtAnkle) = 0 Or g_iLtRed(g_iLtAnkle) = 0 Or g_iLtMM(g_iLtAnkle) = 0) Then
			sError = sError & "The Ankle has not been figured!." & NL
			sError = sError & "Figure the Ankle or use CANCEL to exit" & NL
		End If
		
		
		'Toe style need only be given if a heel tape is given, Heel is at tape 5
		If (g_iFirstTape < 5) And (cboToeStyle.SelectedIndex <= 0) And g_iLegStyle < 3 Then
			sError = sError & "Missing Toe Style." & NL
		End If
		
		'Check that a foot length has been given for self enclosed
		If VB6.GetItemString(cboToeStyle, cboToeStyle.SelectedIndex) = "Self Enclosed" And Val(txtFootLength.Text) = 0 Then
			sError = sError & "A foot length must be given for Self Enclosed toes." & NL
		End If
		
		'Check that a leg style has been chosen
		If g_iLegStyle < 0 Then
			sError = sError & "A Leg Style has not been chosen!." & NL
			sError = sError & "Choose a Style or use CANCEL to exit" & NL
		End If
		
		'General check
		If VB6.GetItemString(cboFabric, cboFabric.SelectedIndex) = "" Then
			sError = sError & "A Fabric has not been chosen!." & NL
		End If
		
		'General check for atemplate
		If cboLeftTemplate.SelectedIndex = -1 Then
			sError = sError & "A template has not been chosen!." & NL
		End If
		
		
		'Display Error message (if required) and return
		If Len(sError) > 0 Then
			MsgBox(sError, 48, "Errors in Data")
			Validate_Data = False
			Exit Function
		Else
			Validate_Data = True
		End If
		
		'For the combo boxes check if a toe selection has been made
		'and a heel tape is not given.
		'This is a none fatal error, the data will be ammended and
		'an information message given
		'will be given
		
		sError = ""
		If nFirstTape > 5 Then
			If cboToeStyle.SelectedIndex > 0 Then
				sError = sError & "Toe style, " & VB6.GetItemString(cboToeStyle, cboToeStyle.SelectedIndex) & NL
			End If
			cboToeStyle.SelectedIndex = 0
		End If
		
		'Display Warning message (if required)
		If Len(sError) > 0 Then
			sError = "The TOE style given can't apply. As there is no Heel." & NL
			sError = sError & "Removing TOE style to reflect this."
			MsgBox(sError, 64, "Warning - Errors in Data")
		End If
		
		
		
		'Yes / No type errors
		'In this case we warn the user that there is a problem!
		'They can continue or they can return to the dialog to make changes
		
		'Initialize error variables
		sError = ""
		iError = False
		
		'Make a stab at seeing if the foot length will be long enough.
		'For Toe = "Self enclosed" or Toe = "Soft enclosed" and a foot length is given
		'This checks that the given foot length is longer than the distance
		'from the heel to the -3 tape at the seam (-1 1/2 for heels less than 9 ")
		'it is a rough giude only.
		If g_nLtLastHeel > 0 And (VB6.GetItemString(cboToeStyle, cboToeStyle.SelectedIndex) = "Self Enclosed" Or (VB6.GetItemString(cboToeStyle, cboToeStyle.SelectedIndex) = "Soft Enclosed" And Val(txtFootLength.Text) <> 0)) Then
			
			'Set checking parameters based on fabric and template
			If g_POWERNET = True Then
				nLargeHeelX = 2.71875
				'Scale heel to the 0 tape of a 30mmHg template
				'This is the worst case for POWERNET fabric
				nHeelLength = g_nLtLastHeel * 0.402
				'Reduce heel by 3
				nHeelLength = nHeelLength + (3 * 0.05025)
			Else
				Select Case cboLeftTemplate.SelectedIndex
					Case 0 ' 13 DS
						nLargeHeelX = 2.73
					Case 1 ' 09 DS
						nLargeHeelX = 2.88
				End Select
				'Scale heel to the 0 tape of a JOBSTEX template
				nHeelLength = g_nLtLastHeel * 0.391
				'Reduce heel by 3
				nHeelLength = nHeelLength + (3 * 0.048875)
			End If
			
			If g_nLtLastHeel >= 9 Then
				nTapeX = nLargeHeelX 'Large heel, distance from heel to -3 tape
			Else
				nTapeX = 1.5 'Small heel, distance from heel to -1 1/2 tape
			End If
			
			nFootLength = ARMEDDIA1.fnDisplaytoInches(Val(txtFootLength.Text))
			
			nValue = System.Math.Sqrt(nHeelLength ^ 2 + nTapeX ^ 2)
			
			If nValue > nFootLength Then
				iError = True
				sError = sError & "The given foot length may be too small to correctly" & NL
				sError = sError & "position the toe!" & NL & NL
				sError = sError & "Do you wish to continue anyway ?"
			End If
		End If
		
		
		'Display Error message (if required) and return
		'These are non-fatal errors
		If iError = True Then
			iError = MsgBox(sError, 36, "Warning, Problems with data")
			If iError = IDYES Then
				Validate_Data = True
			Else
				Validate_Data = False
			End If
		Else
			Validate_Data = True
		End If
		
	End Function
End Class
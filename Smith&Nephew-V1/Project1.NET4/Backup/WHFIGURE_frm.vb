Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class whfigure
	Inherits System.Windows.Forms.Form
	'Project:   WHFIGURE.MAK
	'Form:      WHFIGURE.FRM
	'Purpose:   Dialog to figure the ankles of both legs
	'           for a Waist Height
	'
	'Version:   3.02
	'Date:      16.Jan.95
	'Author:    Gary George
	'-------------------------------------------------------
	'REVISIONS:
	'Date       By      Action
	'-------------------------------------------------------
	'27.Apr.95  GG      Bug Fix, For JOBSTEX_FL the last tape
	'                   was not being found correctly
	'Dec 98     GG      Port to VB5
	'
	'Notes:-
	'
	'MsgBox constant
	'    Const IDYES = 6
	'    Const IDNO = 7
	'    Const IDCANCEL = 2
	
	
	
	
	Private Sub Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Cancel.Click
		Dim Response As Short
		Dim sTask, sCurrentValues As String
		
		'Check if data has been modified
		sCurrentValues = FN_ValuesString()
		
		If sCurrentValues <> g_sChangeChecker Then
			Response = MsgBox("Changes have been made, Save changes before closing", 35, "CAD - Glove Dialogue")
			Select Case Response
				Case IDYES
					PR_CreateMacro_Save("c:\jobst\draw.d")
					sTask = fnGetDrafixWindowTitleText()
					If sTask <> "" Then
						AppActivate(sTask)
						System.Windows.Forms.SendKeys.SendWait("@c:\jobst\draw.d{enter}")
						End
					Else
						MsgBox("Can't find a Drafix Drawing to update!", 16, "WAIST Height - Figure")
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
		PR_FigureRightAnkle()
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
	
	'UPGRADE_WARNING: Event chkRightZipper.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkRightZipper_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkRightZipper.CheckStateChanged
		If g_JOBSTEX = True Or g_JOBSTEX_FL = True Then
			If chkRightZipper.CheckState = 1 Then
				cboRightTemplate.SelectedIndex = 1 '9DS
			Else
				cboRightTemplate.SelectedIndex = 0 '13DS
			End If
		End If
		PR_FigureRightAnkle()
	End Sub
	
	Private Function FN_Open(ByRef sDrafixFile As String, ByRef sType As String, ByRef sName As Object, ByRef sFileNo As Object) As Short
		'Open the DRAFIX macro file
		'Return the file number
		
		'Open file
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = FreeFile
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileOpen(fNum, sDrafixFile, OpenMode.Output)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FN_Open = fNum
		
		'Write header information etc. to the DRAFIX macro file
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//DRAFIX Macro created - " & DateString & "  " & TimeString)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//Patient - " & sName & ", " & sFileNo & "")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//by Visual Basic, Waist Height - Figure")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//type - " & sType & "")
		
		'Clear user selections etc
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "UserSelection (" & QQ & "clear" & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "SetColor" & QC & "Table(" & QQ & "find" & QCQ & "color" & QCQ & "bylayer" & QQ & "));")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "SetStyle" & QC & "Table(" & QQ & "find" & QCQ & "style" & QCQ & "bylayer" & QQ & "));")
		
		'Display Hour Glass symbol
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Display (" & QQ & "cursor" & QCQ & "wait" & QCQ & sType & QQ & ");")
		
		
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
	
	Private Function FN_ValuesString() As String
		'Create a string of all the values given in the
		'Text and Combo boxes.
		'
		'Ignore patient details
		'
		Dim ii As Short
		Dim sString As String
		
		FN_ValuesString = ""
		sString = ""
		
		'Dimensions
		
		FN_ValuesString = sString
		
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
		
		'Disable Time out mechanism
		Timer1.Enabled = False
		Dim nAge, ii, nn, iFabricClass As Short
		Dim iZipper, iMMHg, iValue As Short
		Dim nValue, nStretch As Double
		
		g_sDiagnosis = txtDiagnosis.Text
		
		g_iUidBody = CShort(txtUidBody.Text)
		g_sLastFabric = txtFabric.Text
		
		If ARMEDDIA1.fnGetNumber(txtLegStyle.Text, 1) = 2 Then 'Briefs
			Me.Close()
			'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
			Load(whfigbrf)
			whfigbrf.Show()
			Exit Sub
		End If
		
		'Units
		If txtUnits.Text = "cm" Then
			g_nUnitsFac = 10 / 25.4
		Else
			g_nUnitsFac = 1
		End If
		
		'Establish how the leg is to be figured
		'Base this on chosen fabric and availability of mm
		'If it is not given explicitly
		
		'Check for an explititly given fabric class from the Ankle field
		'This Field will be emplty if no previous figuring has been done
		'Note This is a multi data. I the first number is -ve then there is no ankle.
		'     Also the function fnGetNumber returns -1 if the number does not exist.
		'
		iFabricClass = -1
		If ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 1) <> -1 Then
			iFabricClass = ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 8)
		End If
		If ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 1) <> -1 Then
			iFabricClass = ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 8)
		End If
		
		'Get fabric class from other data
		If iFabricClass < 0 Then
			If txtFabric.Text = "" Then
				If txtDiagnosis.Text = "Burns" Then
					g_POWERNET = True
				Else
					If VB.Left(txtDiagnosis.Text, 5) = "Lymph" Then
						g_JOBSTEX_FL = True
					Else
						g_JOBSTEX = True
					End If
				End If
			Else
				If VB.Left(txtFabric.Text, 3) = "Pow" Then
					g_POWERNET = True
				Else
					If txtLeftMMs.Text <> "" Or txtRightMMs.Text <> "" Then
						g_JOBSTEX_FL = True
					Else
						g_JOBSTEX = True
					End If
				End If
			End If
		Else
			If iFabricClass = 0 Then g_POWERNET = True
			If iFabricClass = 1 Then g_JOBSTEX = True
			If iFabricClass = 2 Then g_JOBSTEX_FL = True
		End If
		
		If g_JOBSTEX = True Then optFabric(1).Checked = True
		If g_JOBSTEX_FL = True Then optFabric(2).Checked = True
		
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
		
		'Update tape boxes
		grdLeftInches.Col = 0
		For ii = 0 To 29
			nValue = Val(Mid(txtLeftLengths.Text, (ii * 4) + 1, 4)) / 10
			If nValue > 0 Then
				txtLeft(ii).Text = CStr(nValue)
				nValue = ARMEDDIA1.fnDisplayToInches(nValue)
				grdLeftInches.Row = ii
				grdLeftInches.Text = ARMEDDIA1.fnInchesToText(nValue)
			End If
		Next ii
		
		grdRightInches.Col = 0
		For ii = 0 To 29
			nValue = Val(Mid(txtRightLengths.Text, (ii * 4) + 1, 4)) / 10
			If nValue > 0 Then
				txtRight(ii).Text = CStr(nValue)
				nValue = ARMEDDIA1.fnDisplayToInches(nValue)
				grdRightInches.Row = ii
				grdRightInches.Text = ARMEDDIA1.fnInchesToText(nValue)
			End If
		Next ii
		
		PR_EstablishAnkles()
		
		If g_POWERNET = True Then PR_EnablePOWERNET()
		If g_JOBSTEX = True Or g_JOBSTEX_FL = True Then PR_EnableJOBSTEX()
		If g_JOBSTEX_FL = True Then PR_EnableFL_JOBSTEX()
		
		'From the ankle tapes extract the saved ankle figuring
		'We have to be careful in the order in which this is set up as in the case
		'of JOBSTEX_FL we do not want to cause recaculation on the ankle.
		'We also must take care only to add values only if they exist
		'(thus we check that the returned value from fnGetNumber is not -ve)
		'
		Dim iFabric As Short
		If txtLeftAnkle.Text <> "" And g_iLtAnkle <> 0 Then
			iZipper = ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 7)
			If iZipper >= 0 Then chkLeftZipper.CheckState = iZipper 'Order is v.important here
			iMMHg = ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 2)
			If iMMHg >= 0 Then
				g_iLtMM(g_iLtAnkle) = iMMHg
				txtLeftMM(g_iLtAnkle).Text = CStr(iMMHg)
			End If
			If g_JOBSTEX_FL = True Then
				iFabric = Val(VB.Left(cboFabric.Text, 2))
				nStretch = ARMEDDIA1.fnGetNumber(txtLeftAnkle.Text, 3)
				If nStretch >= 0 Then
					PR_DisplayFiguredAnkle("Left", nStretch, g_nLtLastAnkle, g_nLtLastHeel, iFabric)
					g_iLtStretch(g_iLtAnkle) = nStretch
				End If
			End If
		End If
		
		If txtRightAnkle.Text <> "" And g_iRtAnkle <> 0 Then
			iZipper = ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 7)
			If iZipper >= 0 Then chkRightZipper.CheckState = iZipper
			iMMHg = ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 2)
			If iMMHg >= 0 Then
				g_iRtMM(g_iRtAnkle) = iMMHg
				txtRightMM(g_iRtAnkle).Text = CStr(iMMHg)
			End If
			If g_JOBSTEX_FL = True Then
				iFabric = Val(VB.Left(cboFabric.Text, 2))
				nStretch = ARMEDDIA1.fnGetNumber(txtRightAnkle.Text, 3)
				If nStretch >= 0 Then
					PR_DisplayFiguredAnkle("Right", nStretch, g_nRtLastAnkle, g_nRtLastHeel, iFabric)
					g_iRtStretch(g_iRtAnkle) = nStretch
				End If
			End If
		End If
		
		'Don't figure ankle for fabric class JOBSTEX_FL as the user can modify the pressures
		'at each leg tape manually and these are stored.  Figuring would overwrite
		'any manual changes
		If g_JOBSTEX_FL <> True Then
			PR_FigureLeftAnkle()
			PR_FigureRightAnkle()
		Else
			If txtRightMMs.Text <> "" And g_iRtAnkle <> 0 Then
				For ii = g_iRtAnkle + 1 To 29
					iValue = Val(Mid(txtRightMMs.Text, (ii * 3) + 1, 3))
					If iValue > 0 Then
						txtRightMM(ii).Text = CStr(iValue)
						g_iRtMM(ii) = iValue
						g_iRtStretch(ii) = Val(Mid(txtRightStr.Text, (ii * 3) + 1, 3))
						g_iRtRed(ii) = Val(Mid(txtRightRed.Text, (ii * 3) + 1, 3))
						grdRightDisplay.Row = ii - 6 'Display from ankle
						grdRightDisplay.Col = 0
						grdRightDisplay.Text = Str(g_iRtStretch(ii))
						grdRightDisplay.Col = 1
						grdRightDisplay.Text = Str(g_iRtRed(ii))
					End If
				Next ii
			End If
			If txtLeftMMs.Text <> "" And g_iLtAnkle <> 0 Then
				For ii = g_iLtAnkle + 1 To 29
					iValue = Val(Mid(txtLeftMMs.Text, (ii * 3) + 1, 3))
					If iValue > 0 Then
						txtLeftMM(ii).Text = CStr(iValue)
						g_iLtMM(ii) = iValue
						g_iLtStretch(ii) = Val(Mid(txtLeftStr.Text, (ii * 3) + 1, 3))
						g_iLtRed(ii) = Val(Mid(txtLeftRed.Text, (ii * 3) + 1, 3))
						grdLeftDisplay.Row = ii - 6 'Display from ankle
						grdLeftDisplay.Col = 0
						grdLeftDisplay.Text = Str(g_iLtStretch(ii))
						grdLeftDisplay.Col = 1
						grdLeftDisplay.Text = Str(g_iLtRed(ii))
					End If
				Next ii
			End If
			
		End If
		
		'Store initial values for use in Cancel_Click
		g_sChangeChecker = FN_ValuesString()
		
		Show()
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default ' Change pointer to default.
		Me.Activate()
		
		
	End Sub
	
	'UPGRADE_ISSUE: Form event Form.LinkExecute was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub Form_LinkExecute(ByRef CmdStr As String, ByRef Cancel As Short)
		If CmdStr = "Cancel" Then
			Cancel = 0
			End
		End If
	End Sub
	
	Private Sub whfigure_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim ii As Short
		Hide()
		'Check if a previous instance is running
		'If it is warn user and exit
		'UPGRADE_ISSUE: App property App.PrevInstance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		If App.PrevInstance Then
			MsgBox("The Figure Module is already running!" & Chr(13) & "Use ALT-TAB and Cancel it.", 16, "Error Starting Figure")
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
		g_JOBSTEX = False
		g_JOBSTEX_FL = False
		g_POWERNET = False
		
		'Initialize globals
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		QQ = Chr(34) 'Double quotes (")
		'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NL = Chr(13) 'New Line
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CC = Chr(44) 'The comma (,)
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
		
		g_sPathJOBST = fnPathJOBST()
		
		'Setup display inches grid
		grdLeftInches.set_ColWidth(0, 700)
		grdRightInches.set_ColWidth(0, 700)
		grdRightInches.set_ColAlignment(0, 2)
		grdLeftInches.set_ColAlignment(0, 2)
		
		For ii = 0 To 29
			grdLeftInches.set_RowHeight(ii, 266)
			grdRightInches.set_RowHeight(ii, 266)
		Next ii
		grdRightDisplay.set_ColWidth(0, 488)
		grdRightDisplay.set_ColWidth(1, 488)
		
		'Setup display of results grid
		For ii = 0 To 1
			grdLeftDisplay.set_ColWidth(ii, 488)
			grdRightDisplay.set_ColWidth(ii, 488)
			grdLeftDisplay.set_ColAlignment(ii, 2)
			grdRightDisplay.set_ColAlignment(ii, 2)
		Next ii
		
		For ii = 0 To 23
			grdLeftDisplay.set_RowHeight(ii, 266)
			grdRightDisplay.set_RowHeight(ii, 266)
		Next ii
		
		'Enable timer
		'This is a time out mechanism that will
		'end the VB programme if the DDE link with DRAFIX
		'is not sucessfully completed
		Timer1.Interval = 6000
		Timer1.Enabled = True
		
	End Sub
	
	Private Sub OK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OK.Click
		'Check that data is all present and insert into drafix
		
		Dim sTask As String
		
		If Validate_Data() Then
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
			Hide()
			
			PR_CreateMacro_Save("c:\jobst\draw.d")
			
			sTask = fnGetDrafixWindowTitleText()
			If sTask <> "" Then
				AppActivate(sTask)
				System.Windows.Forms.SendKeys.SendWait("@C:\JOBST\DRAW.D{enter}")
				'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
				End
			Else
				MsgBox("Can't find a Drafix Drawing to update!", 16, "WAIST Height - Figure")
			End If
		End If
		OK.Enabled = True
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		
	End Sub
	
	'UPGRADE_WARNING: Event optFabric.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optFabric_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optFabric.CheckedChanged
		If eventSender.Checked Then
			Dim index As Short = optFabric.GetIndex(eventSender)
			'Allows the selection of fabric class
			'Used to override the fabric class selection established at link close
			
			'Clean Message box
			labMessage.Text = ""
			
			'Do nothing if same fabric class selected
			If g_POWERNET = True And index = 0 Then Exit Sub
			If g_JOBSTEX = True And index = 1 Then Exit Sub
			If g_JOBSTEX_FL = True And index = 2 Then Exit Sub
			
			'Restablish ankles just in case they have changed
			PR_EstablishAnkles()
			
			'Set the display for the different fabric classes
			'We must first reset the form to the blank form that we use on form load
			PR_ResetFormWHFIGURE()
			
			g_POWERNET = False
			g_JOBSTEX = False
			g_JOBSTEX_FL = False
			
			Select Case index
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
					PR_EnableFL_JOBSTEX()
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
				g_iLtLastFabric = -1 'This will force re-figuring of Left
				g_iRtLastFabric = -1 'This will force re-figuring of Right
				If g_iLtMM(g_iLtAnkle) > 0 Then
					txtLeftMM(g_iLtAnkle).Text = CStr(g_iLtMM(g_iLtAnkle))
					PR_FigureLeftAnkle()
				End If
				If g_iRtMM(g_iRtAnkle) > 0 Then
					txtRightMM(g_iRtAnkle).Text = CStr(g_iRtMM(g_iRtAnkle))
					PR_FigureRightAnkle()
				End If
				g_iLtLastFabric = -1
				g_iRtLastFabric = -1
			End If
			
		End If
	End Sub
	
	Private Sub PR_CreateMacro_Save(ByRef sDrafixFile As String)
		'fNum is a global variable use in subsequent procedures
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = ARMDIA1.FN_Open(sDrafixFile, "Save Data ONLY", (txtPatientName.Text), (txtFileNo.Text))
		
		'If this is a new drawing of a vest then Define the DATA Base
		'fields for the VEST Body and insert the BODYBOX symbol
		BDYUTILS.PR_PutLine("HANDLE hMPD, hBody, hLeftLeg, hRightLeg;")
		
		Update_DDE_Text_Boxes()
		
		Dim sSymbol As String
		
		sSymbol = "waistbody"
		
		If txtUidBody.Text <> "" Then
			'Update the Waist BODY symbol with the fabric
			'UPGRADE_WARNING: Couldn't resolve default property of object QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("hBody = UID (" & QQ & "find" & QC & Val(txtUidBody.Text) & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("if (!hBody) Exit(%cancel," & QQ & "Can't find >" & sSymbol & "< symbol to update!" & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hBody" & CQ & "Fabric" & QCQ & txtFabric.Text & QQ & ");")
		End If
		
		If txtUidRightLeg.Text <> "" Then
			'Update the Right Leg symbol
			'UPGRADE_WARNING: Couldn't resolve default property of object QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("hRightLeg = UID (" & QQ & "find" & QC & Val(txtUidRightLeg.Text) & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("if (!hRightLeg) Exit(%cancel," & QQ & "Can't find >" & sSymbol & "< symbol to update!" & QQ & ");")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hRightLeg" & CQ & "TapeLengthsPt1" & QCQ & Mid(txtRightLengths.Text, 1, 60) & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hRightLeg" & CQ & "TapeLengthsPt2" & QCQ & Mid(txtRightLengths.Text, 61) & QQ & ");")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hRightLeg" & CQ & "TapeMMs" & QCQ & Mid(txtRightMMs.Text, 1, 60) & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hRightLeg" & CQ & "TapeMMs2" & QCQ & Mid(txtRightMMs.Text, 61) & QQ & ");")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hRightLeg" & CQ & "Reduction" & QCQ & Mid(txtRightRed.Text, 1, 60) & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hRightLeg" & CQ & "Reduction2" & QCQ & Mid(txtRightRed.Text, 61) & QQ & ");")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hRightLeg" & CQ & "Grams" & QCQ & Mid(txtRightStr.Text, 1, 60) & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hRightLeg" & CQ & "Grams2" & QCQ & Mid(txtRightStr.Text, 61) & QQ & ");")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hRightLeg" & CQ & "AnkleTape" & QCQ & txtRightAnkle.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hRightLeg" & CQ & "FirstTape" & QCQ & txtRightFirstTape.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hRightLeg" & CQ & "LastTape" & QCQ & txtRightLastTape.Text & QQ & ");")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hRightLeg" & CQ & "Pressure" & QCQ & txtRightTemplate.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hRightLeg" & CQ & "Fabric" & QCQ & txtFabric.Text & QQ & ");")
			
		End If
		
		If txtUidLeftLeg.Text <> "" Then
			'Update Left leg symbol
			'UPGRADE_WARNING: Couldn't resolve default property of object QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("hLeftLeg = UID (" & QQ & "find" & QC & Val(txtUidLeftLeg.Text) & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("if (!hLeftLeg) Exit(%cancel," & QQ & "Can't find >" & sSymbol & "< symbol to update!" & QQ & ");")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hLeftLeg" & CQ & "TapeLengthsPt1" & QCQ & Mid(txtLeftLengths.Text, 1, 60) & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hLeftLeg" & CQ & "TapeLengthsPt2" & QCQ & Mid(txtLeftLengths.Text, 61) & QQ & ");")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hLeftLeg" & CQ & "TapeMMs" & QCQ & Mid(txtLeftMMs.Text, 1, 60) & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hLeftLeg" & CQ & "TapeMMs2" & QCQ & Mid(txtLeftMMs.Text, 61) & QQ & ");")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hLeftLeg" & CQ & "Reduction" & QCQ & Mid(txtLeftRed.Text, 1, 60) & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hLeftLeg" & CQ & "Reduction2" & QCQ & Mid(txtLeftRed.Text, 61) & QQ & ");")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hLeftLeg" & CQ & "Grams" & QCQ & Mid(txtLeftStr.Text, 1, 60) & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hLeftLeg" & CQ & "Grams2" & QCQ & Mid(txtLeftStr.Text, 61) & QQ & ");")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hLeftLeg" & CQ & "AnkleTape" & QCQ & txtLeftAnkle.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hLeftLeg" & CQ & "FirstTape" & QCQ & txtLeftFirstTape.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hLeftLeg" & CQ & "LastTape" & QCQ & txtLeftLastTape.Text & QQ & ");")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hLeftLeg" & CQ & "Pressure" & QCQ & txtLeftTemplate.Text & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BDYUTILS.PR_PutLine("SetDBData( hLeftLeg" & CQ & "Fabric" & QCQ & txtFabric.Text & QQ & ");")
			
		End If
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(fNum)
		
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
			'Ankle selection
			Case 900
				sError = sError & "As one ankle is at the +1½ tape and the other is at the +3 tape. "
				sError = sError & "Figuring for both Ankles at the +3 tape."
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
			'UPGRADE_WARNING: Couldn't resolve default property of object sContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			labMessage.Text = sContext + NL + sError
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object sContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			labMessage.Text = labMessage.Text & NL & NL + sContext + NL + sError
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
		
		If sLeg = "Left" Then
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
		Else
			grdRightDisplay.Row = g_iRtAnkle - 6
			grdRightDisplay.Col = 0
			'Stretch
			grdRightDisplay.Text = Str(nStretch)
			'Reduction
			grdRightDisplay.Col = 1
			grdRightDisplay.Text = Str(iReduction)
			g_iRtRed(g_iRtAnkle) = iReduction
			'Max Stretch
			labRightMaxStr.Text = CStr(iDonningStretch)
		End If
		
		'Set Max and Min stretches
		iMaxStretch = 44
		If iFabric <= 55 Then iMinStretch = 16 Else iMinStretch = 18
		
		If nStretch < iMinStretch Then
			PR_DisplayErrorMessage(2002, sLeg & " Ankle Figuring")
		End If
		
		If nStretch > iMaxStretch Then
			PR_DisplayErrorMessage(2001, sLeg & " Ankle Figuring")
		End If
		
		If iDonningStretch > 110 Then
			PR_DisplayErrorMessage(2003, sLeg & " Ankle Figuring")
		End If
		
		If iReduction > 32 Then
			PR_DisplayErrorMessage(2004, sLeg & " Ankle Figuring")
		End If
		
		
		
	End Sub
	
	Private Sub PR_EnableFL_JOBSTEX()
		'Setup the fields to allow calculation of
		'stretch at each tape
		Dim ii, iValue As Short
		
		'Right Leg
		If g_iRtAnkle <> 0 Then
			For ii = 8 To 29
				txtRightMM(ii).Visible = True
			Next ii
			txtRightMM(g_iRtAnkle + 1).Visible = True
			
			'Load MM data from string
			'Ankle data will be stored seperatly
			If txtRightMMs.Text <> "" Then
				For ii = g_iRtAnkle + 1 To 29
					iValue = Val(Mid(txtRightMMs.Text, (ii * 3) + 1, 3))
					If iValue > 0 Then
						txtRightMM(ii).Text = CStr(iValue)
					End If
				Next ii
			End If
			
		End If
		
		'Left Leg
		If g_iLtAnkle <> 0 Then
			'Make Boxes visible
			For ii = 8 To 29
				txtLeftMM(ii).Visible = True
			Next ii
			
			txtLeftMM(g_iLtAnkle + 1).Visible = True
			
			'Load MM data from string
			'Ankle data will be stored seperatly
			If txtLeftMMs.Text <> "" Then
				For ii = g_iLtAnkle + 1 To 29
					iValue = Val(Mid(txtLeftMMs.Text, (ii * 3) + 1, 3))
					If iValue > 0 Then
						txtLeftMM(ii).Text = CStr(iValue)
					End If
				Next ii
			End If
		End If
		
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
		
		cboRightTemplate.Items.Clear()
		cboRightTemplate.Items.Add("13 DS")
		cboRightTemplate.Items.Add("09 DS")
		cboRightTemplate.Enabled = False
		
		
		'Set value
		For ii = 0 To (cboLeftTemplate.Items.Count - 1)
			If VB6.GetItemString(cboLeftTemplate, ii) = txtLeftTemplate.Text Then
				cboLeftTemplate.SelectedIndex = ii
			End If
		Next ii
		
		For ii = 0 To (cboRightTemplate.Items.Count - 1)
			If VB6.GetItemString(cboRightTemplate, ii) = txtRightTemplate.Text Then
				cboRightTemplate.SelectedIndex = ii
			End If
		Next ii
		
		'Set the display for each leg depending on the ankle and fabric and options.
		'g_JOBSTEX_FL is a special case where all of the stretch at tapes above the
		'ankle are calculated
		'
		'Panty Legs
		'If a template is not given then
		'then set 13DS for JOBSTEX, this is first item in list
		'
		
		If g_iRtAnkle = 0 Then
			If cboRightTemplate.SelectedIndex = -1 Then cboRightTemplate.SelectedIndex = 0
			cboRightTemplate.Enabled = False
			chkRightZipper.Enabled = False
			chkRightZipper.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
		If g_iLtAnkle = 0 Then
			If cboLeftTemplate.SelectedIndex = -1 Then cboLeftTemplate.SelectedIndex = 0
			cboLeftTemplate.Enabled = False
			chkLeftZipper.Enabled = False
			chkLeftZipper.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
		
		If g_iRtAnkle <> 0 Then
			'Set display headings
			labRightMaxStr.Visible = True
			labRightDisp(1).Text = "Str"
			For ii = 0 To 3
				labRightDisp(ii).Visible = True
			Next 
			
			'Set display value depending on Ankle position
			iEnable = g_iRtAnkle
			If iEnable = 6 Then
				iDisable = 7
			Else
				iDisable = 6
			End If
			
			txtRightMM(iEnable).Visible = True
			txtRightMM(iDisable).Visible = False
			
		End If
		
		
		If g_iLtAnkle <> 0 Then
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
		
		cboRightTemplate.Items.Clear()
		cboRightTemplate.Items.Add("30mm Hg")
		cboRightTemplate.Items.Add("35mm Hg")
		cboRightTemplate.Items.Add("40mm Hg")
		cboRightTemplate.Items.Add("50mm Hg")
		
		cboLeftTemplate.Enabled = True
		cboRightTemplate.Enabled = True
		
		'Set value
		For ii = 0 To (cboLeftTemplate.Items.Count - 1)
			If VB6.GetItemString(cboLeftTemplate, ii) = txtLeftTemplate.Text Then
				cboLeftTemplate.SelectedIndex = ii
			End If
		Next ii
		
		For ii = 0 To (cboRightTemplate.Items.Count - 1)
			If VB6.GetItemString(cboRightTemplate, ii) = txtRightTemplate.Text Then
				cboRightTemplate.SelectedIndex = ii
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
		
		If g_iRtAnkle = 0 Then
			If cboRightTemplate.SelectedIndex = -1 Then cboRightTemplate.SelectedIndex = 0
			chkRightZipper.Enabled = False
			chkRightZipper.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
		If g_iLtAnkle = 0 Then
			If cboLeftTemplate.SelectedIndex = -1 Then cboLeftTemplate.SelectedIndex = 0
			chkLeftZipper.Enabled = False
			chkLeftZipper.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
		If g_iRtAnkle <> 0 Then
			labRightDisp(1).Text = "Gms"
			labRightDisp(3).Visible = False
			labRightMaxStr.Visible = False
			For ii = 0 To 2
				labRightDisp(ii).Visible = True
			Next 
			
			'Set display value depending on Ankle position
			iEnable = g_iRtAnkle
			If iEnable = 6 Then
				iDisable = 7
			Else
				iDisable = 6
			End If
			
			txtRightMM(iEnable).Visible = True
			txtRightMM(iDisable).Visible = False
		End If
		
		If g_iLtAnkle <> 0 Then
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
		g_nLtLastHeel = ARMEDDIA1.fnDisplayToInches(Val(txtLeft(5).Text))
		If g_nLtLastHeel <> 0 Then
			If g_nLtLastHeel < 9 Then
				g_iLtAnkle = 6 'Small heel Ankle at +1 1/2 tape
			Else
				g_iLtAnkle = 7 'Large heel Ankle at +3 tape
			End If
		End If
		
		g_iRtAnkle = 0
		g_nRtLastAnkle = 0
		g_nRtLastHeel = ARMEDDIA1.fnDisplayToInches(Val(txtRight(5).Text))
		If g_nRtLastHeel <> 0 Then
			If g_nRtLastHeel < 9 Then
				g_iRtAnkle = 6 'Small heel Ankle at +1 1/2 tape
			Else
				g_iRtAnkle = 7 'Large heel Ankle at +3 tape
			End If
		End If
		
		'Check for asymetric ankles
		If g_iRtAnkle <> g_iLtAnkle And g_iRtAnkle <> 0 And g_iLtAnkle <> 0 Then
			g_iLtAnkle = 7 'Large heel Ankle at +3 tape
			g_iRtAnkle = 7 'Large heel Ankle at +3 tape
			PR_DisplayErrorMessage(900, "Ankle selection")
		End If
		
		g_nLtLastAnkle = ARMEDDIA1.fnDisplayToInches(Val(txtLeft(g_iLtAnkle).Text))
		g_nRtLastAnkle = ARMEDDIA1.fnDisplayToInches(Val(txtRight(g_iRtAnkle).Text))
		
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
		Dim nPatternDim As Double
		
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
		'    The POWERNET fabric is based on the file g_sPathJOBST + "\WHFABRIC.DAT"
		'    For clarity this file can contain blank lines.  Therefor we must check
		'    that the fabric selected is not blank.
		If VB6.GetItemString(cboFabric, cboFabric.SelectedIndex) = "" Then Exit Sub
		
		'Establish if any of the inputs have changed compared to the last stored
		'values.
		'If they have changed then recalculate and redisplay
		'NB  Use of GOTO, (structured programmers get stuffed!)
		
		'Heel Tape
		If ARMEDDIA1.fnDisplayToInches(Val(txtLeft(5).Text)) <> g_nLtLastHeel Then GoTo CALC
		
		'Ankle Tape
		If ARMEDDIA1.fnDisplayToInches(Val(txtLeft(g_iLtAnkle).Text)) <> g_nLtLastAnkle Then GoTo CALC
		
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
		
		nHeelCir = ARMEDDIA1.fnDisplayToInches(Val(txtLeft(5).Text))
		g_nLtLastHeel = nHeelCir
		
		nAnkleCir = ARMEDDIA1.fnDisplayToInches(Val(txtLeft(g_iLtAnkle).Text))
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
			nLastTape = ii - 1
			
			'Establish pressure at last tape
			'As the last tape is at a 20 reduction we know that the stretch must
			'be 25%
			nLastTapeCir = ARMEDDIA1.fnDisplayToInches(Val(txtLeft(nLastTape).Text))
			nLastTapeMMHg = ARMDIA1.round(FN_JOBSTEX_Pressure(nLastTapeCir, nHeelCir, 25, iZipper, iFabric))
			
			txtLeftMM(nLastTape).Text = CStr(nLastTapeMMHg)
			g_iLtStretch(nLastTape) = 25
			g_iLtRed(nLastTape) = 20
			g_iLtMM(nLastTape) = nLastTapeMMHg
			
			'First pass from Ankle to LastTape
			nStretch = g_iLtLastStretch
			iPrevMMHg = g_iLtLastMM
			For ii = g_iLtAnkle + 1 To nLastTape - 1
				nTapeCir = ARMEDDIA1.fnDisplayToInches(Val(txtLeft(ii).Text))
				nMMHg = 1000 'Force While
				While nMMHg > iPrevMMHg And nStretch > 1
					nMMHg = ARMDIA1.round(FN_JOBSTEX_Pressure(nTapeCir, nHeelCir, nStretch, iZipper, iFabric))
					If nMMHg > iPrevMMHg Then nStretch = nStretch - 1
				End While
				txtLeftMM(ii).Text = CStr(nMMHg)
				g_iLtMM(ii) = nMMHg
				iPrevMMHg = nMMHg
				
				'Revise stretch to that based on the Pressure
				nStretch = ARMDIA1.round(FN_JOBSTEX_Stretch(nTapeCir, nHeelCir, nMMHg, iZipper, iFabric))
				
				g_iLtStretch(ii) = nStretch
				g_iLtRed(ii) = ARMDIA1.round((1 - 1 / (0.01 * nStretch + 1)) * 100)
			Next ii
			
			'Second pass from LastTape to Ankle
			'Based on last tape Pressure recalculate the tapes until the pressures
			'become the same
			iPrevMMHg = nLastTapeMMHg
			
			For ii = nLastTape - 1 To g_iLtAnkle + 1 Step -1
				nTapeCir = ARMEDDIA1.fnDisplayToInches(Val(txtLeft(ii).Text))
				
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
				
				'Revise stretch to that based on the Pressure
				nStretch = ARMDIA1.round(FN_JOBSTEX_Stretch(nTapeCir, nHeelCir, nMMHg, iZipper, iFabric))
				
				g_iLtStretch(ii) = nStretch
				g_iLtRed(ii) = ARMDIA1.round((1 - 1 / ((0.01 * nStretch) + 1)) * 100)
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
	
	Private Sub PR_FigureLeftTape(ByRef index As Short)
		'Figures the Pressure/Stretch at a LEFT LegTape
		'This is only available to JOBSTEX_FL
		'See PR_FigureLeftTape for details
		
		Dim nMMHg, nStretch As Double
		Dim nTapeCir As Double
		Dim iZipper, iFabric As Short
		
		'Establish if enough data exists to calculate.
		'If not then EXIT quietly
		If Val(txtLeftMM(index).Text) <= 0 Then Exit Sub
		
		'Establish if pressure has changed from last time
		'If it hasn't then exit
		'This is used to cure a problem in that the stretch for a tape is based not on the
		'pressure but on the previous or subsequent stretch, depending if it was found
		'on the forward pass or the backward pass.
		'The pressure was then back calculated fron the stretch.
		'We had a problem in that the same pressue can give a different stretch
		'To stop this happening when tabbing through the pressures this check has
		'been introduced
		'
		If Val(txtLeftMM(index).Text) = g_iLtMM(index) Then Exit Sub
		
		'MMHg and Stretch for JOBSTEX fabric
		'Don't Calculate if no values for Ankle tape
		If (g_JOBSTEX = True Or g_JOBSTEX_FL = True) And Val(txtLeftMM(g_iLtAnkle).Text) = 0 Or g_iLtStretch(g_iLtAnkle) = 0 Then Exit Sub
		
		'If we get to here we can assume that there is enough data to process and
		'we can calculate the relevant values.
		
		nTapeCir = ARMEDDIA1.fnDisplayToInches(Val(txtLeft(index).Text))
		iZipper = chkLeftZipper.CheckState
		iFabric = Val(VB.Left(VB6.GetItemString(cboFabric, cboFabric.SelectedIndex), 2))
		nMMHg = Val(txtLeftMM(index).Text)
		g_iLtMM(index) = nMMHg
		
		'Calculate stretch
		nStretch = FN_JOBSTEX_Stretch(nTapeCir, g_nLtLastHeel, nMMHg, iZipper, iFabric)
		
		'Store values
		g_iLtStretch(index) = ARMDIA1.round(nStretch)
		g_iLtRed(index) = ARMDIA1.round((1 - 1 / (0.01 * nStretch + 1)) * 100)
		
		'Display values
		'Note, the display grid displays from the ankle only
		'
		grdLeftDisplay.Row = index - 6 'Display from ankle
		grdLeftDisplay.Col = 0
		grdLeftDisplay.Text = Str(g_iLtStretch(index))
		grdLeftDisplay.Col = 1
		grdLeftDisplay.Text = Str(g_iLtRed(index))
		
	End Sub
	
	Private Sub PR_FigureRightAnkle()
		'Figures the Pressure/Stretch at the Right ankle
		'Input from form:-
		'   Right(5-7) where     Right(5) = Heel Tape
		'                       Right(6) = Ankle Tape (Heel < 9")
		'                       Right(7) = Ankle Tape (Heel > 9")
		'   cboFabric
		'   chkRightZipper
		'   txtRightMM(g_iRtAnkle)
		'   txtRightStretch(g_iRtAnkle)   JOBSTEX fabric only
		'
		'Input from globals
		'   g_JOBSTEX           Flag to indicate use of JOBSTEX fabric
		'   g_JOBSTEX_FL        Flag to indicate that JOBSTEX fabric in use
		'                       and that the stretch will be calculated at all
		'                       leg tape
		'   g_POWERNET          Flag to indicate use of Powernet fabric
		'
		'   g_iRtAnkle          Indicates the Right ankle tape
		'   g_nRtLastHeel
		'   g_nRtLastAnkle
		'   g_iRtLastMM
		'   g_iRtLastStretch    JOBSTEX fabric only
		'   g_iRtLastZipper
		'   g_iLastFabric
		'
		'JOBSTEX fabric Input and Output
		'   txtRightMM(g_iRtAnkle)        User Input, if txtRightStretch has been modified.
		'   txtRightStretch(g_iRtAnkle)   User Input, if txtRightMM has been modified.
		'   labRightRed(g_iRtAnkle)       Displays Reduction
		'   labRightMaxStr(2)             Displays MaxStretch
		'
		'Note:-
		'   The user inputs a required pressure in MMHg the stretch is then calculated
		'   and the stretch dependant values displayed.
		'   ARternately the stretch can be entered and the pressure in MMHg calculated.
		'   Given input of both (eg at link close) then MMHg takes precidence
		'
		'POWERNET fabric Input and Output
		'   txtRightMM(g_iRtAnkle)            User Input
		'   txtRightStretch(g_iRtAnkle)       not used (disabled)
		'   labRightFiguredGrams(g_iRtAnkle)  Displays Grams
		'   labRightRed(g_iRtAnkle)           Displays Reduction
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
		If Val(txtRight(5).Text) = 0 Or Val(txtRight(g_iRtAnkle).Text) = 0 Or cboFabric.SelectedIndex = -1 Then
			Exit Sub
		End If
		
		'MMHg for POWERNET fabric
		If g_POWERNET = True And Val(txtRightMM(g_iRtAnkle).Text) = 0 Then Exit Sub
		
		'MMHg and Stretch for JOBSTEX fabric
		If (g_JOBSTEX = True Or g_JOBSTEX_FL = True) And Val(txtRightMM(g_iRtAnkle).Text) = 0 Then Exit Sub
		
		'Fabric Selection
		'NOTE:-
		'    The POWERNET fabric is based on the file g_sPathJOBST + "\WHFABRIC.DAT"
		'    For clarity this file can contain blank lines.  Therefor we must check
		'    that the fabric selected is not blank.
		If VB6.GetItemString(cboFabric, cboFabric.SelectedIndex) = "" Then Exit Sub
		
		'Establish if any of the inputs have changed compared to the last stored
		'values.
		'If they have changed then recalculate and redisplay
		'NB  Use of GOTO, (structured programmers get stuffed!)
		
		'Heel Tape
		If ARMEDDIA1.fnDisplayToInches(Val(txtRight(5).Text)) <> g_nRtLastHeel Then GoTo RT_CALC
		
		'Ankle Tape
		If ARMEDDIA1.fnDisplayToInches(Val(txtRight(g_iRtAnkle).Text)) <> g_nRtLastAnkle Then GoTo RT_CALC
		
		'Pressure MMHg
		If Val(txtRightMM(g_iRtAnkle).Text) <> g_iRtLastMM Then GoTo RT_CALC
		
		'Zippers
		If chkRightZipper.CheckState <> g_iRtLastZipper Then GoTo RT_CALC
		
		'Fabric
		If cboFabric.SelectedIndex <> g_iRtLastFabric Then GoTo RT_CALC
		
		'If nothing has changed then Exit this sub
		Exit Sub
		
		'If we get to here we can assume that there is enough data to process and
		'we can calculate the relevant values.
RT_CALC: 
		'Calculate stretch based on given pressure (MMHg).
		'NB  All dimensions are converted to in inches except that last dimesions
		'    are stored using their display values.
		
		nHeelCir = ARMEDDIA1.fnDisplayToInches(Val(txtRight(5).Text))
		g_nRtLastHeel = nHeelCir
		
		nAnkleCir = ARMEDDIA1.fnDisplayToInches(Val(txtRight(g_iRtAnkle).Text))
		g_nRtLastAnkle = nAnkleCir
		
		iZipper = chkRightZipper.CheckState
		g_iRtLastZipper = iZipper
		
		g_iRtLastFabric = cboFabric.SelectedIndex
		nLastTape = -1
		
		nMMHg = Val(txtRightMM(g_iRtAnkle).Text)
		If g_JOBSTEX = True Or g_JOBSTEX_FL = True Then
			'JOBSTEX - Fabric
			iFabric = Val(VB.Left(VB6.GetItemString(cboFabric, cboFabric.SelectedIndex), 2))
			nStretch = ARMDIA1.round(FN_JOBSTEX_Stretch(nAnkleCir, nHeelCir, nMMHg, iZipper, iFabric))
			
			'Calculate and Display the other stretch dependent values
			'This procedure also issues warnings a well
			PR_DisplayFiguredAnkle("Right", nStretch, nAnkleCir, nHeelCir, iFabric)
			
			g_iRtLastStretch = nStretch
			'g_iRtRed(g_iRtAnkle) this is set in PR_DisplayFiguredAnkle above
			g_iRtStretch(g_iRtAnkle) = nStretch
			g_iRtMM(g_iRtAnkle) = nMMHg
			
			
			'Check if template is set
			If cboRightTemplate.SelectedIndex = -1 Then cboRightTemplate.SelectedIndex = iZipper
			
		Else
			'POWERNET - Fabric
			iGrams = ARMDIA1.round(nAnkleCir * nMMHg)
			iModulus = Val(Mid(VB6.GetItemString(cboFabric, cboFabric.SelectedIndex), 5, 3))
			iReduction = FN_POWERNET_Reduction(iGrams, iModulus)
			'Exit if error
			If iReduction < 0 Then
				PR_DisplayErrorMessage(iReduction, "Error - Right Ankle Figuring")
				Exit Sub
			End If
			
			'Minimum Reduction for all diagnosis
			If iReduction < 14 Then
				iReduction = 14
				nMMHg = FN_POWERNET_Pressure(iReduction, nAnkleCir, iModulus)
				iGrams = ARMDIA1.round(nAnkleCir * nMMHg)
				PR_DisplayErrorMessage(1001, "Right Ankle Figuring")
			End If
			
			'Maximum Reduction for Burns
			If iReduction > 26 And txtDiagnosis.Text = "Burns" Then
				iReduction = 26
				nMMHg = FN_POWERNET_Pressure(iReduction, nAnkleCir, iModulus)
				iGrams = ARMDIA1.round(nAnkleCir * nMMHg)
				PR_DisplayErrorMessage(1002, "Right Ankle Figuring")
			End If
			
			'Maximum reduction for all other diagnosis
			If iReduction > 32 Then
				iReduction = 32
				nMMHg = FN_POWERNET_Pressure(iReduction, nAnkleCir, iModulus)
				iGrams = ARMDIA1.round(nAnkleCir * nMMHg)
				PR_DisplayErrorMessage(1003, "Right Ankle Figuring")
			End If
			
			'Display Calculations
			txtRightMM(g_iRtAnkle).Text = Str(nMMHg)
			grdRightDisplay.Row = g_iRtAnkle - 6
			grdRightDisplay.Col = 0
			grdRightDisplay.Text = Str(iGrams)
			grdRightDisplay.Col = 1
			grdRightDisplay.Text = Str(iReduction)
			
			'Set Template
			Select Case iReduction
				Case 0 To 14, 14 To 18
					cboRightTemplate.SelectedIndex = 0 '30 mmhg
				Case 19 To 21
					cboRightTemplate.SelectedIndex = 1 '35 mmhg
				Case 21 To 23
					cboRightTemplate.SelectedIndex = 2 '40 mmhg
				Case 24 To 32, Is > 32
					cboRightTemplate.SelectedIndex = 3 '50 mmhg
			End Select
			
			'Max Stretch (Only if no zipper given)
			If chkRightZipper.CheckState = 0 Then
				nHeelCir = nHeelCir / 4
				nAnkleCir = (nAnkleCir * ((100 - iReduction) / 100)) / 2
				If nAnkleCir < nHeelCir Then PR_DisplayErrorMessage(1004, "Right Ankle Figuring")
			End If
			
			g_iRtStretch(g_iRtAnkle) = iGrams 'Bit of a misnomer but what the Hell!
			g_iRtRed(g_iRtAnkle) = iReduction
			g_iRtMM(g_iRtAnkle) = nMMHg
			
		End If
		g_iRtLastMM = Val(txtRightMM(g_iRtAnkle).Text)
		
		'Calculate the pressure at all leg tapes
		'
		If g_JOBSTEX_FL = True Then
			'iZipper = 0     'Don't use zipper
			
			'Establish Last tape position
			'This allows for the addition of leg tapes to extend up the leg
			'It's not very efficient but what the hell!!!!
			nLastTape = 0
			For ii = g_iRtAnkle + 1 To 29
				If txtRight(ii).Text = "" Then Exit For
			Next ii
			nLastTape = ii - 1
			
			'Establish pressure at last tape
			'As the last tape is at a 20 reduction we know that the stretch must
			'be 25%
			nLastTapeCir = ARMEDDIA1.fnDisplayToInches(Val(txtRight(nLastTape).Text))
			nLastTapeMMHg = ARMDIA1.round(FN_JOBSTEX_Pressure(nLastTapeCir, nHeelCir, 25, iZipper, iFabric))
			
			txtRightMM(nLastTape).Text = CStr(nLastTapeMMHg)
			g_iRtStretch(nLastTape) = 25
			g_iRtRed(nLastTape) = 20
			g_iRtMM(nLastTape) = nLastTapeMMHg
			
			'First pass from Ankle to LastTape
			nStretch = g_iRtLastStretch
			iPrevMMHg = g_iRtLastMM
			For ii = g_iRtAnkle + 1 To nLastTape - 1
				nTapeCir = ARMEDDIA1.fnDisplayToInches(Val(txtRight(ii).Text))
				nMMHg = 1000 'Force While
				While nMMHg > iPrevMMHg And nStretch > 1
					nMMHg = ARMDIA1.round(FN_JOBSTEX_Pressure(nTapeCir, nHeelCir, nStretch, iZipper, iFabric))
					If nMMHg > iPrevMMHg Then nStretch = nStretch - 1
				End While
				txtRightMM(ii).Text = CStr(nMMHg)
				g_iRtMM(ii) = nMMHg
				iPrevMMHg = nMMHg
				
				'Revise stretch to that based on the Pressure
				nStretch = ARMDIA1.round(FN_JOBSTEX_Stretch(nTapeCir, nHeelCir, nMMHg, iZipper, iFabric))
				
				g_iRtStretch(ii) = nStretch
				g_iRtRed(ii) = ARMDIA1.round((1 - 1 / (0.01 * nStretch + 1)) * 100)
			Next ii
			
			'Second pass from LastTape to Ankle
			'Based on last tape Pressure recalculate the tapes until the pressures
			'become the same
			iPrevMMHg = nLastTapeMMHg
			
			For ii = nLastTape - 1 To g_iRtAnkle + 1 Step -1
				nTapeCir = ARMEDDIA1.fnDisplayToInches(Val(txtRight(ii).Text))
				
				nMMHg = Val(txtRightMM(ii).Text)
				nStretch = g_iRtStretch(ii)
				If nMMHg >= iPrevMMHg Then Exit For
				While nMMHg < iPrevMMHg
					nMMHg = ARMDIA1.round(FN_JOBSTEX_Pressure(nTapeCir, nHeelCir, nStretch, iZipper, iFabric))
					If nMMHg < iPrevMMHg Then nStretch = nStretch + 1
				End While
				txtRightMM(ii).Text = CStr(nMMHg)
				g_iRtMM(ii) = nMMHg
				iPrevMMHg = nMMHg
				
				'Revise stretch to that based on the Pressure
				nStretch = ARMDIA1.round(FN_JOBSTEX_Stretch(nTapeCir, nHeelCir, nMMHg, iZipper, iFabric))
				
				g_iRtStretch(ii) = nStretch
				g_iRtRed(ii) = ARMDIA1.round((1 - 1 / ((0.01 * nStretch) + 1)) * 100)
			Next ii
			
			'Display Values
			'STRETCH
			grdRightDisplay.Col = 0
			For ii = g_iRtAnkle + 1 To nLastTape
				grdRightDisplay.Row = ii - 6
				grdRightDisplay.Text = Str(g_iRtStretch(ii))
			Next ii
			
			'REDUCTION
			grdRightDisplay.Col = 1
			For ii = g_iRtAnkle + 1 To nLastTape
				grdRightDisplay.Row = ii - 6
				grdRightDisplay.Text = Str(g_iRtRed(ii))
			Next ii
			
		End If
		
	End Sub
	
	Private Sub PR_FigureRightTape(ByRef index As Short)
		Dim nMMHg, nStretch As Double
		Dim nTapeCir As Double
		Dim iZipper, iFabric As Short
		
		'Establish if enough data exists to calculate.
		'If not then EXIT quietly
		If Val(txtRightMM(index).Text) <= 0 Then Exit Sub
		
		If Val(txtRightMM(index).Text) = g_iRtMM(index) Then Exit Sub
		
		'MMHg and Stretch for JOBSTEX fabric
		'Don't Calculate if no values for Ankle tape
		If (g_JOBSTEX = True Or g_JOBSTEX_FL = True) And Val(txtRightMM(g_iRtAnkle).Text) = 0 Or g_iRtStretch(g_iRtAnkle) = 0 Then Exit Sub
		
		'If we get to here we can assume that there is enough data to process and
		'we can calculate the relevant values.
		
		nTapeCir = ARMEDDIA1.fnDisplayToInches(Val(txtRight(index).Text))
		iZipper = chkRightZipper.CheckState
		iFabric = Val(VB.Left(VB6.GetItemString(cboFabric, cboFabric.SelectedIndex), 2))
		nMMHg = Val(txtRightMM(index).Text)
		g_iRtMM(index) = nMMHg
		
		'Calculate stretch
		nStretch = FN_JOBSTEX_Stretch(nTapeCir, g_nRtLastHeel, nMMHg, iZipper, iFabric)
		
		'Store values
		g_iRtStretch(index) = ARMDIA1.round(nStretch)
		g_iRtRed(index) = ARMDIA1.round((1 - 1 / (0.01 * nStretch + 1)) * 100)
		
		'Display values
		'Note, the display grid displays from the ankle only
		'
		grdRightDisplay.Row = index - 6 'Display from ankle
		grdRightDisplay.Col = 0
		grdRightDisplay.Text = Str(g_iRtStretch(index))
		grdRightDisplay.Col = 1
		grdRightDisplay.Text = Str(g_iRtRed(index))
		
	End Sub
	
	Private Sub PR_HeelChange(ByRef sLeg As String)
		'Procedure to modify the display if heel changed
		
		Dim iLtExistingAnkle, iRtExistingAnkle As Short
		'Clean Message box
		labMessage.Text = ""
		
		'Exit if heel has not been changed
		If sLeg = "Left" Then
			If ARMEDDIA1.fnDisplayToInches(Val(txtLeft(5).Text)) = g_nLtLastHeel Then Exit Sub
		Else
			If ARMEDDIA1.fnDisplayToInches(Val(txtRight(5).Text)) = g_nRtLastHeel Then Exit Sub
		End If
		
		'Store existing ankle position and get new ones
		iLtExistingAnkle = g_iLtAnkle
		iRtExistingAnkle = g_iRtAnkle
		
		PR_EstablishAnkles()
		
		'If revised the same as existing then exit
		If iLtExistingAnkle = g_iLtAnkle And iRtExistingAnkle = g_iRtAnkle Then Exit Sub
		
		'Modify display
		PR_ResetFormWHFIGURE()
		If g_POWERNET = True Then PR_EnablePOWERNET()
		If g_JOBSTEX = True Or g_JOBSTEX_FL = True Then PR_EnableJOBSTEX()
		If g_JOBSTEX_FL = True Then PR_EnableFL_JOBSTEX()
		
	End Sub
	
	Private Sub PR_LoadFabricFromFile(ByRef sFileName As String)
		'Procedure to load the POWERNET conversion chart from file
		'N.B. File opening Errors etc. are not handled (so tough titty!)
		
		Dim fNum, ii As Short
		fNum = FreeFile
		FileOpen(fNum, sFileName, OpenMode.Input)
		ii = 0
		Do Until EOF(fNum)
			Input(fNum, POWERNET.Modulus(ii))
			Input(fNum, POWERNET.Conversion_Renamed(ii))
			ii = ii + 1
		Loop 
		FileClose(fNum)
		
	End Sub
	
	Private Sub PR_PutLine(ByRef sLine As String)
		'Puts the contents of sLine to the opened "Macro" file
		'Puts the line with no translation or additions
		'    fNum is global variable
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, sLine)
		
	End Sub
	
	Private Sub PR_ResetFormWHFIGURE()
		'Reset display of form to a clean state
		Dim ii As Short
		
		'Reset Max Stretch
		labLeftMaxStr.Text = ""
		labLeftMaxStr.Visible = False
		labLeftDisp(3).Visible = False
		labRightMaxStr.Text = ""
		labRightMaxStr.Visible = False
		labRightDisp(3).Visible = False
		
		'Clean Display grid
		For ii = 0 To 23
			
			grdLeftDisplay.Row = ii
			grdLeftDisplay.Col = 0
			grdLeftDisplay.Text = ""
			grdLeftDisplay.Col = 1
			grdLeftDisplay.Text = ""
			
			grdRightDisplay.Row = ii
			grdRightDisplay.Col = 0
			grdRightDisplay.Text = ""
			grdRightDisplay.Col = 1
			grdRightDisplay.Text = ""
			
		Next ii
		
		'Switch off MM text boxes
		For ii = 6 To 29
			txtLeftMM(ii).Text = ""
			txtLeftMM(ii).Visible = False
			txtRightMM(ii).Text = ""
			txtRightMM(ii).Visible = False
		Next ii
		
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		'If this event happens then the DDE link to
		'DRAFIX has not been terminated sucessfully.
		'IE No link close.
		'We assume a failure of some sort and we End
		
		End
		
	End Sub
	
	Private Sub txtLeft_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLeft.Enter
		Dim index As Short = txtLeft.GetIndex(eventSender)
		ARMEDDIA1.Select_Text_In_Box(txtLeft(index))
	End Sub
	
	Private Sub txtLeft_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLeft.Leave
		Dim index As Short = txtLeft.GetIndex(eventSender)
		BDLEGDIA1.Validate_And_Display_Text_In_Box(txtLeft(index), grdLeftInches, index)
		If index = g_iLtAnkle Then PR_FigureLeftAnkle()
		If index = 5 Then PR_HeelChange("Left")
	End Sub
	
	Private Sub txtLeftMM_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLeftMM.Leave
		Dim index As Short = txtLeftMM.GetIndex(eventSender)
		If index = g_iLtAnkle Then
			PR_FigureLeftAnkle()
		Else
			PR_FigureLeftTape(index)
		End If
	End Sub
	
	Private Sub txtRight_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRight.Enter
		Dim index As Short = txtRight.GetIndex(eventSender)
		ARMEDDIA1.Select_Text_In_Box(txtRight(index))
	End Sub
	
	Private Sub txtRight_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRight.Leave
		Dim index As Short = txtRight.GetIndex(eventSender)
		BDLEGDIA1.Validate_And_Display_Text_In_Box(txtRight(index), grdRightInches, index)
		If index = g_iRtAnkle Then PR_FigureRightAnkle()
		If index = 5 Then PR_HeelChange("Right")
	End Sub
	
	Private Sub txtRightMM_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRightMM.Leave
		Dim index As Short = txtRightMM.GetIndex(eventSender)
		If index = g_iRtAnkle Then
			PR_FigureRightAnkle()
		Else
			PR_FigureRightTape(index)
		End If
		
	End Sub
	
	Private Sub Update_DDE_Text_Boxes()
		'Called from OK_Click
		'Update the text boxes used for DDE transfer
		Dim ii As Short
		Dim sString, sJustifiedString As String
		Dim sFabricClass As String
		Dim iStr, iRed, iMM As Short
		Dim sStr, sRed, sMM As String
		Dim nValue As Single
		
		If g_POWERNET = True Then sFabricClass = "0"
		If g_JOBSTEX = True Then sFabricClass = "1"
		If g_JOBSTEX_FL = True Then sFabricClass = "2"
		
		txtFabric.Text = VB6.GetItemString(cboFabric, cboFabric.SelectedIndex)
		
		'LEFT Leg
		'Assume that data has been validated earlier
		
		'Set initial values
		txtLeftLengths.Text = ""
		txtLeftRed.Text = ""
		txtLeftStr.Text = ""
		txtLeftMMs.Text = ""
		txtLeftFirstTape.Text = CStr(-1)
		txtLeftLastTape.Text = CStr(-1)
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
			If CDbl(txtLeftFirstTape.Text) < 0 And nValue > 0 Then txtLeftFirstTape.Text = CStr(ii + 1)
			If CDbl(txtLeftLastTape.Text) < 0 And CDbl(txtLeftFirstTape.Text) > 0 And nValue = 0 Then txtLeftLastTape.Text = CStr(ii)
		Next ii
		
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
		
		'Transfer to the "legbox" the following data in blank delimited format
		'
		'txtLeftAnkleTape = "AnkleTape#
		'                    Pressure
		'                    [Grams|Stretch]
		'                    Reduction
		'                    AnkleLength
		'                    HeelLength
		'                    Zipper Status
		'                    FabricClass
		'
		'This means that the CutOut drawing Macro (WH_CUT.D) need not load all
		'of the routines to read leg data, as it is only interested in the
		'heel length
		'
		'g_iLtAnkle=0 => footless, so don't worry about the ankle
		'Note:-
		'    The DRAFIX macros index tape lengths from 1 not 0 as VB does.
		'    Therefor g_iLtAnkle is incremented by 1
		'
		sString = ""
		If g_iLtAnkle <> 0 Then
			sString = sString & Str(g_iLtAnkle + 1) & " "
			sString = sString & Str(g_iLtMM(g_iLtAnkle)) & " "
			sString = sString & Str(g_iLtStretch(g_iLtAnkle)) & " "
			sString = sString & Str(g_iLtRed(g_iLtAnkle)) & " "
			sString = sString & Str(g_nLtLastAnkle) & " "
			sString = sString & Str(g_nLtLastHeel) & " "
			sString = sString & Str(g_iLtLastZipper) & " "
			txtLeftAnkle.Text = sString & sFabricClass
		Else
			'Set the txtAnkleTape to reflect the footless state
			txtLeftAnkle.Text = "-1 0 0 0 0 0 0 " & sFabricClass
		End If
		
		
		'RIGHT Leg
		
		txtRightLengths.Text = ""
		txtRightRed.Text = ""
		txtRightStr.Text = ""
		txtRightMMs.Text = ""
		txtRightFirstTape.Text = CStr(-1)
		txtRightLastTape.Text = CStr(-1)
		txtRightTemplate.Text = VB6.GetItemString(cboRightTemplate, cboRightTemplate.SelectedIndex)
		
		For ii = 0 To 29
			nValue = Val(txtRight(ii).Text)
			If nValue <> 0 Then
				nValue = nValue * 10 'Shift decimal place to right
				sJustifiedString = New String(" ", 4)
				sJustifiedString = RSet(Trim(Str(nValue)), Len(sJustifiedString))
			Else
				sJustifiedString = New String(" ", 4)
			End If
			
			'Tape values
			txtRightLengths.Text = txtRightLengths.Text & sJustifiedString
			
			'Set first and last tape (assumes no holes in data)
			If CDbl(txtRightFirstTape.Text) < 0 And nValue > 0 Then txtRightFirstTape.Text = CStr(ii + 1)
			If CDbl(txtRightLastTape.Text) < 0 And CDbl(txtRightFirstTape.Text) > 0 And nValue = 0 Then txtRightLastTape.Text = CStr(ii)
		Next ii
		
		'Where JOBSTEX_FL has been choosen then we update the MMs, Reduction and Stretch
		'These will then be picked up by the Leg Drawing modules
		If g_JOBSTEX_FL = True And g_iRtAnkle <> 0 Then
			For ii = 0 To 29
				iMM = g_iRtMM(ii)
				iRed = g_iRtRed(ii)
				iStr = g_iRtStretch(ii)
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
				txtRightMMs.Text = txtRightMMs.Text & sMM
				txtRightRed.Text = txtRightRed.Text & sRed
				txtRightStr.Text = txtRightStr.Text & sStr
			Next ii
		End If
		
		'Transfer to the "legbox" the following data in blank delimited format
		
		sString = ""
		If g_iRtAnkle <> 0 Then
			sString = sString & Str(g_iRtAnkle + 1) & " "
			sString = sString & Str(g_iRtMM(g_iRtAnkle)) & " "
			sString = sString & Str(g_iRtStretch(g_iRtAnkle)) & " "
			sString = sString & Str(g_iRtRed(g_iRtAnkle)) & " "
			sString = sString & Str(g_nRtLastAnkle) & " "
			sString = sString & Str(g_nRtLastHeel) & " "
			sString = sString & Str(g_iRtLastZipper) & " "
			txtRightAnkle.Text = sString & sFabricClass
		Else
			'Set the txtAnkleTape to reflect the footless state
			txtRightAnkle.Text = "-1 0 0 0 0 0 0 " & sFabricClass
		End If
		
	End Sub
	
	Private Function Validate_Data() As Short
		'Called from OK_Click
		
		Dim NL, sError, sTextList As String
		Dim sLeftError, sRightError As String
		Dim nFirstTape, ii, nLastTape As Short
		Dim iError As Short
		Dim nFootLength, nHeelLength, nTapeX As Double
		Dim nValue As Double
		
		NL = Chr(10) 'new line
		sTextList = "  7  6 4½  3 1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"
		
		
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
				sError = sError & "Missing Tape length - " & LTrim(Mid(sTextList, (ii * 3) + 1, 3)) & NL
			End If
		Next ii
		
		'Check if -3 tape exists for Large Heel
		If g_iLtAnkle = 7 And Val(txtLeft(3).Text) = 0 Then
			sError = sError & "As the Heel is 9 inches and over." & NL
			sError = sError & "and there is no -3 tape the foot will not draw properly " & NL
		End If
		
		'Check if -1 1/2 tape exists for Small Heel
		If g_iLtAnkle = 6 And Val(txtLeft(4).Text) = 0 Then
			sError = sError & "As the Heel is smaller than 9 inches." & NL
			sError = sError & "and there is no -1 1/2 tape the foot will not draw properly " & NL
		End If
		
		'Check that figuring has been done at the ankle
		If g_iLtAnkle <> 0 And (g_iLtStretch(g_iLtAnkle) = 0 Or g_iLtRed(g_iLtAnkle) = 0 Or g_iLtMM(g_iLtAnkle) = 0) Then
			sError = sError & "The Ankle has not been figured!." & NL
			sError = sError & "Figure the Ankle or use CANCEL to exit" & NL
		End If
		
		sLeftError = sError
		
		'RIGHT LEG Checks
		'Check Tape length data text boxes for holes (ie missing values)
		'Establish First and last tape
		
		sError = ""
		
		For ii = 0 To 29 Step 1
			If Val(txtRight(ii).Text) > 0 Then Exit For
		Next ii
		nFirstTape = ii
		
		For ii = 29 To 0 Step -1
			If Val(txtRight(ii).Text) > 0 Then Exit For
		Next ii
		nLastTape = ii
		
		For ii = nFirstTape To nLastTape
			If Val(txtRight(ii).Text) = 0 Then
				sError = sError & "Missing Tape length - " & LTrim(Mid(sTextList, (ii * 3) + 1, 3)) & NL
			End If
		Next ii
		
		'Check if -3 tape exists for Large Heel
		If g_iRtAnkle = 7 And Val(txtRight(3).Text) = 0 Then
			sError = sError & "As the Heel is 9 inches and over." & NL
			sError = sError & "and there is no -3 tape the foot will not draw properly " & NL
		End If
		
		'Check if -1 1/2 tape exists for Small Heel
		If g_iRtAnkle = 6 And Val(txtRight(4).Text) = 0 Then
			sError = sError & "As the Heel is smaller than 9 inches." & NL
			sError = sError & "and there is no -1 1/2 tape the foot will not draw properly " & NL
		End If
		
		'Check that figuring has been done at the ankle
		If g_iRtAnkle <> 0 And (g_iRtStretch(g_iRtAnkle) = 0 Or g_iRtRed(g_iRtAnkle) = 0 Or g_iRtMM(g_iRtAnkle) = 0) Then
			sError = sError & "The Ankle has not been figured!." & NL
			sError = sError & "Figure the Ankle or use CANCEL to exit" & NL
		End If
		
		sRightError = sError
		
		If sRightError <> "" Then sError = "----------------------------- RIGHT LEG -----------------------------" & NL & sRightError
		If sLeftError <> "" Then sError = sError & "----------------------------- LEFT  LEG -----------------------------" & NL & sLeftError
		
		'General check
		If VB6.GetItemString(cboFabric, cboFabric.SelectedIndex) = "" Then
			sError = sError & "A Fabric has not been chosen!." & NL
		End If
		
		'Display Error message (if required) and return
		If Len(sError) > 0 Then
			MsgBox(sError, 48, "Errors in Data")
			Validate_Data = False
			Exit Function
		Else
			Validate_Data = True
		End If
		
		
		
		'Yes / No type errors
		'In this case we warn the user that there is a problem!
		'They can continue or they can return to the dialog to make changes
		
		'Initialize error variables
		sError = ""
		iError = False
		
		'Display Error message (if required) and return
		'These are non-fatal errors
		If iError = True Then
			iError = MsgBox(sError, 36, "Warning, Problems with data")
			If iError = 6 Then 'IDYES
				Validate_Data = True
			Else
				Validate_Data = False
			End If
		Else
			Validate_Data = True
		End If
		
	End Function
End Class
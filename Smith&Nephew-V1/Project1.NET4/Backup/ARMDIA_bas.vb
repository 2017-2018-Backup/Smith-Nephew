Option Strict Off
Option Explicit On
Module ARMDIA1
	'XY data type to represent points
	Structure XY
		Dim X As Double
		Dim y As Double
	End Structure
	
	Structure curve
		Dim n As Short
		<VBFixedArray(100)> Dim X() As Double
		<VBFixedArray(100)> Dim y() As Double
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			'UPGRADE_WARNING: Lower bound of array X was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim X(100)
			'UPGRADE_WARNING: Lower bound of array y was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim y(100)
		End Sub
	End Structure
	
	'PI as a Global Constant
	Public Const PI As Double = 3.141592654
	
	Public MainForm As armdia
	
	'Globals set by FN_Open
	Public CC As Object 'Comma
	Public QQ As Object 'Quote
	Public NL As Object 'Newline
	Public fNum As Object 'Macro file number
	Public QCQ As Object 'Quote Comma Quote
	Public QC As Object 'Quote Comma
	Public CQ As Object 'Comma Quote
	
	'Constants to define the drive and the root directory
	Public g_sFileNo As String 'The patients file no
	Public g_sSide As String 'The side Left or right
	Public g_sPatient As String 'The patients name
	Public g_sFlapLength As String
	Public g_sFabric As String
	Public g_sFabnam As String
	Public g_sPressureChange As String
	Public g_sFabricChange As String
	Public g_sWpleat1 As String
	Public g_sWpleat2 As String
	Public g_sSpleat1 As String
	Public g_sSpleat2 As String
	Public g_sFlapStrap As String
	Public g_sFlapChk As Object
	Public g_sinchflag As String
	Public g_sFlap As String
	Public g_sMM As String
	Public g_sGaunt As String
	Public g_sDetGaunt As String
	Public g_sNoThumb As String
	Public g_sPalmNo As String
	Public g_sWristNo As String
	Public g_sPalmWristDist As String
	Public g_sThumbCircum As String
	Public g_sThumbLength As String
	Public g_sEnclosedThumb As String
	Public g_sSecondLastTape As String
	Public g_sSecondTape As String
	Public g_sFirstTape As String
	Public g_sLastTape As String
	Public g_sModulus As String
	Public g_sStump As String
	Public g_sTapeLengths As String
	Public g_sTapeMMs As String
	Public g_sGrams As String
	Public g_sReduction As String
	Public g_sAmm As Object
	Public g_sBmm As Object
	Public g_sCmm As Object
	Public g_sDmm As Object
	Public g_sWaistCir As Object
	
	Public g_sWorkOrder As String
	
	'Store current layer and text setings to reduce DRAFIX code
	'this value is checked in PR_SetLayer
	Public g_sCurrentLayer As String
	Public g_nCurrTextHt As Object
	Public g_nCurrTextAspect As Object
	Public g_nCurrTextHorizJust As Object
	Public g_nCurrTextVertJust As Object
	Public g_nCurrTextFont As Object
	Public g_nCurrTextAngle As Object
	
	Public g_sHoleCheck As String
	Public g_sOld As String
	
	Public g_sID As String
	Public g_sEditLengths As String
	Public g_sUnits As String
	
	
	Public g_iStyleFirstTape As Short
	Public g_iStyleLastTape As Short
	Public g_iFirstTape As Short
	Public g_iLastTape As Short
	Public g_nFrontStrapLength As Double
	Public g_nGauntletExtension As Double
	
	Public g_Modified As Short
	Public g_sPathJOBST As String
	
	Public g_bDrawVestMesh As Short
	Public g_bDrawBodyMesh As Short
	
	Public Detachable, WristNo, PalmNo, ThumLen, EnclosedThm As Object
	
	
	
	
	Sub DataBaseDataUpDate(ByRef sType As String)
		Dim sSymbol As Object
		Dim sString As String
		Dim sGauntlet, sFlap As String
		Dim sBoxType As String
		Dim nYoffset As Single
		
		
		'Choose data box depending on type
		If armdia.txtType.Text = "ARM" Then
			sBoxType = "armarm"
		Else
			sBoxType = "vestarm"
		End If
		
		'Get sleevebox handle
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Close(" & QQ & "selection" & QC & "hChan);")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hChan = Open(" & QQ & "selection" & QCQ & "DB SymbolName = '" & sBoxType & "' AND DB Sleeve ='" & g_sSide & "'" & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "ResetSelection(hChan);")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hSleeve = GetNextSelection(hChan);")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Close(" & QQ & "selection" & QC & "hChan);")
		
		'If sleevebox not found then insert one
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "if(!hSleeve){")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "    Close(" & QQ & "selection" & QC & "hChan);")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "    hChan = Open(" & QQ & "selection" & QCQ & "DB SymbolName = 'mainpatientdetails'" & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "    ResetSelection(hChan);")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "    hTitle = GetNextSelection(hChan);")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "    Close(" & QQ & "selection" & QC & "hChan);")
		'Get title box origin
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "    GetGeometry(hTitle,&sTitleName, &xyTitleOrigin,&xyTitleScale,&aTitleAngle);")
		
		'Insert arm or sleeve box
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "    if ( !Symbol(" & QQ & "find" & QCQ & sBoxType & QQ & ")) Exit(%cancel," & QQ & "Cant find SLEEVEBOX or ARMBOX symbol to insert|nCheck your installation, that JOBST.SLB exists" & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "    if(StringCompare(" & QQ & "Left" & QCQ & g_sSide & QQ & ")) xyTitleOrigin.x = xyTitleOrigin.x + 1.5;")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "        else xyTitleOrigin.x = xyTitleOrigin.x + 3;")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "    Execute(" & QQ & "menu" & QCQ & "SetLayer" & QC & "Table(" & QQ & "find" & QCQ & "layer" & QCQ & "Data" & QQ & "));")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "    hSleeve = AddEntity(" & QQ & "symbol" & QCQ & sBoxType & QC & "xyTitleOrigin);")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "}")
		
		'Insert Arm common
		'UPGRADE_WARNING: Couldn't resolve default property of object sSymbol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sSymbol = "armcommon"
		If armdia.txtUidAC.Text = "" Then
			'Insert a new symbol
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "if ( Symbol(" & QQ & "find" & QCQ & sSymbol & QQ & ")){")
			'        Print #fNum, "if(!hSleeve){"
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "  Close(" & QQ & "selection" & QC & "hChan);")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "  hChan = Open(" & QQ & "selection" & QCQ & "DB SymbolName = 'mainpatientdetails'" & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "  ResetSelection(hChan);")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "  hTitle = GetNextSelection(hChan);")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "  Close(" & QQ & "selection" & QC & "hChan);")
			'Get title box origin
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "  GetGeometry(hTitle,&sTitleName, &xyTitleOrigin,&xyTitleScale,&aTitleAngle);")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "  Execute (" & QQ & "menu" & QCQ & "SetLayer" & QC & "Table(" & QQ & "find" & QCQ & "layer" & QCQ & "Data" & QQ & "));")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "  hSym = AddEntity(" & QQ & "symbol" & QCQ & sSymbol & QC & "xyTitleOrigin.x, xyTitleOrigin.y);")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "  }")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "else")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "  Exit(%cancel, " & QQ & "Can't find >" & sSymbol & "< symbol to insert\nCheck your installation, that JOBST.SLB exists!" & QQ & ");")
		Else
			'Use existing symbol
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "hSym = UID (" & QQ & "find" & QC & Val(armdia.txtUidAC.Text) & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "if (!hSym) Exit(%cancel," & QQ & "Can't find >" & sSymbol & "< symbol to update!" & QQ & ");")
		End If
		'Update DB fields
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData( hSym" & CQ & "Fabric" & QCQ & armdia.txtFabric.Text & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData( hSym" & CQ & "fileno" & QCQ & armdia.txtFileNo.Text & QQ & ");")
		
		
		
		
		
		
		'Flap multiple field
		sString = New String(" ", 35)
		sString = LSet(g_sFlap, Len(sString))
		If Val(armdia.txtCustFlapLength.Text) = 0 Then armdia.txtCustFlapLength.Text = CStr(-1)
		If Val(armdia.txtWaistCir.Text) = 0 Then armdia.txtWaistCir.Text = "-1"
		If Val(armdia.txtFrontStrapLength.Text) = 0 Then armdia.txtFrontStrapLength.Text = "-1"
		sFlap = sString & armdia.txtStrap.Text & " " & armdia.txtCustFlapLength.Text & " " & armdia.txtWaistCir.Text & " " & armdia.txtFrontStrapLength.Text
		
		'Gauntlet Multiple field
		If g_sGaunt = "False" Then
			sGauntlet = "0"
		Else
			'Set falgs
			sString = "1 "
			If g_sEnclosedThumb = "True" Then
				sString = sString & "1 "
			Else
				sString = sString & "0 "
			End If
			
			If g_sDetGaunt = "True" Then
				sString = sString & "1 "
			Else
				sString = sString & "0 "
			End If
			
			If g_sNoThumb = "True" Then
				sString = sString & "1 "
			Else
				sString = sString & "0 "
			End If
			If Val(armdia.txtGauntletExtension.Text) = 0 Then armdia.txtGauntletExtension.Text = "-1"
			'Data
			sGauntlet = sString & Val(g_sWristNo) & " " & Val(g_sPalmNo) & " " & Val(g_sThumbLength) & " " & Val(g_sThumbCircum) & " " & Val(g_sPalmWristDist) & " " & armdia.txtGauntletExtension.Text
		End If
		
		'Update data base fields
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData(hSleeve," & QQ & "Sleeve" & QCQ & g_sSide & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData(hSleeve," & QQ & "Fileno" & QCQ & g_sFileNo & QQ & ");")
		
		'Wrist and Shoulder pleat multiple fields
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData(hSleeve," & QQ & "WristPleat" & QCQ & g_sWpleat1 & " " & g_sWpleat2 & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData(hSleeve," & QQ & "ShoulderPleat" & QCQ & g_sSpleat1 & " " & g_sSpleat2 & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData(hSleeve," & QQ & "Flap" & QCQ & sFlap & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData(hSleeve," & QQ & "Gauntlet" & QCQ & sGauntlet & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData(hSleeve," & QQ & "MM" & QCQ & g_sMM & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData(hSleeve," & QQ & "Modulus" & QCQ & g_sModulus & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData(hSleeve," & QQ & "Fabric" & QCQ & g_sFabric & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData(hSleeve," & QQ & "Stump" & QCQ & g_sStump & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData(hSleeve," & QQ & "TapeMMs" & QCQ & g_sTapeMMs & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData(hSleeve," & QQ & "Reduction" & QCQ & g_sReduction & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData(hSleeve," & QQ & "Grams" & QCQ & g_sGrams & QQ & ");")
		
		'Tape lengths are universal, Store the actually used lengths with the
		'Profile Origin Marker
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData(hSleeve," & QQ & "TapeLengths" & QCQ & g_sTapeLengths & QQ & ");")
		'Store ID of last drawn
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetDBData(hSleeve," & QQ & "ID" & QCQ & g_sID & QQ & ");")
		
		
		If sType = "Draw" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData(hOrigin," & QQ & "Fabric" & QCQ & g_sFabric & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData(hOrigin," & QQ & "Gauntlet" & QCQ & sGauntlet & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData(hOrigin," & QQ & "TapeLengths" & QCQ & g_sEditLengths & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData(hOrigin," & QQ & "units" & QCQ & g_sUnits & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData(hOrigin," & QQ & "WristPleat" & QCQ & g_sWpleat1 & " " & g_sWpleat2 & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData(hOrigin," & QQ & "ShoulderPleat" & QCQ & g_sSpleat1 & " " & g_sSpleat2 & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData(hOrigin," & QQ & "Flap" & QCQ & sFlap & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData(hOrigin," & QQ & "Grams" & QCQ & g_sGrams & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData(hOrigin," & QQ & "Reduction" & QCQ & g_sReduction & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData(hOrigin," & QQ & "TapeMMs" & QCQ & g_sTapeMMs & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData(hOrigin," & QQ & "Stump" & QCQ & g_sStump & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData(hOrigin," & QQ & "Modulus" & QCQ & g_sModulus & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData(hOrigin," & QQ & "MM" & QCQ & g_sMM & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetDBData(hOrigin," & QQ & "age" & QCQ & CType(armdia.Controls("txtAge"), Object).Text & QQ & ");")
		End If
		
	End Sub
	
	
	
	Function FN_CalcAngle(ByRef xyStart As XY, ByRef xyEnd As XY) As Double
		'Function to return the angle between two points in degrees
		'in the range 0 - 360
		'Zero is always 0 and is never 360
		
		Dim X, y As Object
		Dim rAngle As Double
		
		'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		X = xyEnd.X - xyStart.X
		'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		y = xyEnd.y - xyStart.y
		
		'Horizomtal
		'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If X = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If y > 0 Then
				FN_CalcAngle = 90
			Else
				FN_CalcAngle = 270
			End If
			Exit Function
		End If
		
		'Vertical (avoid divide by zero later)
		'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If y = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If X > 0 Then
				FN_CalcAngle = 0
			Else
				FN_CalcAngle = 180
			End If
			Exit Function
		End If
		
		'All other cases
		'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rAngle = System.Math.Atan(y / X) * (180 / PI) 'Convert to degrees
		
		If rAngle < 0 Then rAngle = rAngle + 180 'rAngle range is -PI/2 & +PI/2
		
		'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If y > 0 Then
			FN_CalcAngle = rAngle
		Else
			FN_CalcAngle = rAngle + 180
		End If
		
	End Function
	
	Function FN_CalcLength(ByRef xyStart As XY, ByRef xyEnd As XY) As Double
		'Fuction to return the length between two points
		'Greatfull thanks to Pythagorus
		
		FN_CalcLength = System.Math.Sqrt((xyEnd.X - xyStart.X) ^ 2 + (xyEnd.y - xyStart.y) ^ 2)
		
	End Function
	
	Function FN_EscapeQuotesInString(ByRef sAssignedString As Object) As String
		'Search through the string looking for " (double quote characater)
		'If found use \ (Backslash) to escape it
		'
		Dim ii As Short
		'UPGRADE_NOTE: Char was upgraded to Char_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Char_Renamed As String
		Dim sEscapedString As String
		
		FN_EscapeQuotesInString = ""
		
		For ii = 1 To Len(sAssignedString)
			'UPGRADE_WARNING: Couldn't resolve default property of object sAssignedString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Char_Renamed = Mid(sAssignedString, ii, 1)
			If Char_Renamed = """" Then
				sEscapedString = sEscapedString & "\" & Char_Renamed
			Else
				sEscapedString = sEscapedString & Char_Renamed
			End If
		Next ii
		
		FN_EscapeQuotesInString = sEscapedString
		
	End Function
	
	Function FN_EscapeSlashesInString(ByRef sAssignedString As Object) As String
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
	
	Function FN_GetNumber(ByVal sString As String, ByRef iIndex As Short) As Double
		'Function to return as a numerical value the iIndexth item in a string
		'that uses blanks (spaces) as delimiters.
		'EG
		'    sString = "12.3 65.1 45"
		'    FN_GetNumber( sString, 2) = 65.1
		'
		'If the iIndexth item is not found then return -1 to indicate an error.
		'This assumes that the string will not be used to store -ve numbers.
		'Indexing starts from 1
		
		Dim ii, iPos As Short
		Dim sItem As String
		
		'Initial error checking
		sString = Trim(sString) 'Remove leading and trailing blanks
		
		If Len(sString) = 0 Then
			FN_GetNumber = -1
			Exit Function
		End If
		
		'Prepare string
		sString = sString & " " 'Trailing blank as stopper for last item
		
		'Get iIndexth item
		For ii = 1 To iIndex
			iPos = InStr(sString, " ")
			If ii = iIndex Then
				sString = Left(sString, iPos - 1)
				FN_GetNumber = Val(sString)
				Exit Function
			Else
				sString = LTrim(Mid(sString, iPos))
				If Len(sString) = 0 Then
					FN_GetNumber = -1
					Exit Function
				End If
			End If
		Next ii
		
		'The function should have exited befor this, however just in case
		'(iIndex = 0) we indicate an error,
		FN_GetNumber = -1
		
	End Function
	
	Function FN_Open(ByRef sDrafixFile As String, ByRef sName As Object, ByRef sPatientFile As Object, ByRef sLeftorRight As Object, ByRef sType As String) As Short
		'Open the DRAFIX macro file
		'Initialise Global variables
		Dim sID, sAxillaType As String
		Dim iPos As Double
		Dim sProximalStyle, sStyle, sDistalStyle As String
		
		'Open file
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fNum = FreeFile
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileOpen(fNum, sDrafixFile, OpenMode.Output)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FN_Open = fNum
		
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
		
		'Initialise patient globals
		'UPGRADE_WARNING: Couldn't resolve default property of object sPatientFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_sFileNo = sPatientFile
		'UPGRADE_WARNING: Couldn't resolve default property of object sLeftorRight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_sSide = sLeftorRight
		'UPGRADE_WARNING: Couldn't resolve default property of object sName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_sPatient = sName
		If CType(armdia.Controls("txtWorkOrder"), Object).Text = "" Then
			g_sWorkOrder = "-"
		Else
			g_sWorkOrder = CType(armdia.Controls("txtWorkOrder"), Object).Text
		End If
		
		
		'Globals to reduced drafix code written to file
		g_sCurrentLayer = ""
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextHt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextHt = 0.125
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextAspect = 0.6
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextHorizJust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextHorizJust = 1 'Left
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextVertJust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextVertJust = 32 'Bottom
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextFont. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextFont = 0 'BLOCK
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g_nCurrTextAngle = 0
		
		'Create the 4 character string to identify the type
		PR_GetStyle(sStyle)
		
		
		'Write header information etc. to the DRAFIX macro file
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//DRAFIX Macro created - " & DateString & "  " & TimeString)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//Patient - " & g_sPatient & CC & " " & g_sFileNo & CC & " Sleeve-" & g_sSide)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "//by Visual Basic")
		
		'Define DRAFIX variables
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "HANDLE hLayer, hSleeve, hTitle, hChan, hEnt, hOrigin, hSym;")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "XY     xyAxilla, xyAxillaLow, xyTitleOrigin, xyStart, xyTitleScale, xyOrigin, xyElbow;")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "ANGLE  aTitleAngle;")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "STRING sTitleName, sFileNo, sSleeve, sID, sWorkOrder, sName, sVestID, sData, sDate, sPathJOBST;")
		
		'Set path to JOBST installed directory
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "sPathJOBST = " & QQ & FN_EscapeSlashesInString(g_sPathJOBST) & QQ & ";")
		
		'Set up ID  data base field
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Table(" & QQ & "add" & QCQ & "field" & QCQ & "ID" & QCQ & "string" & QQ & ");")
		
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
		
		'Clear user selections etc
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "UserSelection (" & QQ & "clear" & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Execute (" & QQ & "menu" & QCQ & "SetStyle" & QC & "Table(" & QQ & "find" & QCQ & "style" & QCQ & "bylayer" & QQ & "));")
		
		g_sID = sStyle & g_sFileNo & g_sSide
		
		'Find axilla type
		If CType(armdia.Controls("txtType"), Object).Text <> "ARM" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			iPos = InStr(CType(armdia.Controls("txtVestRaglan"), Object).Text, QQ) 'Regular axilla are awkward as they
			'contain a space, look for quote (")
			'QQ is a constant set to quote (")
			If iPos <> 0 Then
				'Escape quote for use in drafix
				'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sAxillaType = Left(CType(armdia.Controls("txtVestRaglan"), Object).Text, iPos - 1) & "\" & QQ
			Else
				iPos = InStr(CType(armdia.Controls("txtVestRaglan"), Object).Text, " ")
				sAxillaType = Left(CType(armdia.Controls("txtVestRaglan"), Object).Text, iPos - 1)
			End If
			
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "sSleeve = " & QQ & g_sSide & QQ & ";")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "sFileNo = " & QQ & g_sFileNo & QQ & ";")
		
		'Get Start point
		If sType = "Draw" Then
			'Get Start point
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetSymbolLibrary( sPathJOBST + " & QQ & "\\JOBST.SLB" & QQ & ");")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "GetUser (" & QQ & "xy" & QCQ & "Indicate Start Point" & QC & "&xyStart);")
			
			'Place a marker at the start point for later use.
			'Get a UID and create the unique 4 character start to the ID code
			'Note this is a bit dogey if the drawing contains more than 9999 entities
			PR_SetLayer("Construct")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "hOrigin = AddEntity(" & QQ & "marker" & QCQ & "xmarker" & QC & "xyStart" & CC & "0.125);")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "if (hOrigin) {")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "  sData=StringMiddle(MakeString(" & QQ & "long" & QQ & ",UID(" & QQ & "get" & QQ & ",hOrigin)), 1, 4) ; ")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "  while (StringLength(sData) < 4) sData = sData + " & QQ & " " & QQ & ";")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "  sData = sData + sFileNo + sSleeve ;")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "  SetDBData( hOrigin," & QQ & "ID" & QQ & CC & QQ & g_sID & "originmark" & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "  SetDBData( hOrigin," & QQ & "Data" & QQ & CC & "sData" & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "  SetDBData( hOrigin," & QQ & "curvetype" & QQ & CC & QQ & "sleeveoriginmark" & QQ & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "  }")
		End If
		
		'Display Hour Glass symbol
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Display (" & QQ & "cursor" & QCQ & "wait" & QCQ & "Drawing" & QQ & ");")
		
		'Set values for use futher on by other macros
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "xyOrigin = xyStart" & ";")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "sID = " & QQ & g_sID & QQ & ";")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "sName = " & QQ & g_sPatient & QQ & ";")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "sWorkOrder = " & QQ & armdia.txtWorkOrder.Text & QQ & ";")
		
	End Function
	
	Function fnGetString(ByVal sString As String, ByRef iIndex As Short, ByRef sDelimiter As String) As String
		'Function to return as a string the iIndexth item in a string
		'that using the given string sDelimiter as the delimiter.
		'EG
		'    sString = "Sam Spade Hello"
		'    sDelimiter = " " {SPACE}
		'    fnGetNumber( sString, 2) = "Spade"
		'
		'If the iIndexth item is not found then return "" to indicate an error.
		'Indexing starts from 1
		
		Dim ii, iPos As Short
		Dim sItem As String
		
		'Initial error checking
		sString = Trim(sString) 'Remove leading and trailing blanks
		
		If Len(sString) = 0 Then
			fnGetString = ""
			Exit Function
		End If
		
		'Prepare string
		sString = sString & sDelimiter 'Trailing sDelimiter as stopper for last item
		
		'Get iIndexth item
		For ii = 1 To iIndex
			iPos = InStr(sString, sDelimiter)
			If ii = iIndex Then
				sString = Left(sString, iPos - 1)
				fnGetString = sString
				Exit Function
			Else
				sString = LTrim(Mid(sString, iPos + 1))
				If Len(sString) = 0 Then
					fnGetString = ""
					Exit Function
				End If
			End If
		Next ii
		
		'The function should have exited befor this, however just in case
		'(iIndex = 0) we indicate an error,
		fnGetString = ""
		
	End Function
	
	Function max(ByRef nFirst As Object, ByRef nSecond As Object) As Object
		' Returns the maximum of two numbers
		'UPGRADE_WARNING: Couldn't resolve default property of object nSecond. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object nFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If nFirst >= nSecond Then
			'UPGRADE_WARNING: Couldn't resolve default property of object nFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object max. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			max = nFirst
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object nSecond. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object max. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			max = nSecond
		End If
	End Function
	
	Sub PR_AddEntityID(ByRef sFileNo As String, ByRef sSide As String, ByRef sType As Object)
		'To the DRAFIX macro file (given by the global fNum)
		'write the syntax to add to an ENTITY the database information
		'in the DB variable "ID" that will allow the identity of an entity
		'to be retrieved, by other parts of the system.
		'
		'For this to work it assumes that the following DRAFIX variables
		'are defined.
		'    HANDLE  hEnt
		'
		'Note:-
		'    fNum, CC, QQ, NL are globals initialised by FN_Open
		'
		Dim sID As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object sType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sID = sFileNo & sSide & sType
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "if (hEnt) SetDBData( hEnt," & QQ & "ID" & QQ & CC & QQ & sID & QQ & ");")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "if (hEnt) SetDBData( hEnt," & QQ & "Data" & QQ & CC & "sData" & ");")
		
	End Sub
	
	Sub PR_CalcPolar(ByRef xyStart As XY, ByVal nAngle As Double, ByRef nLength As Double, ByRef xyReturn As XY)
		'Procedure to return a point at a distance and an angle from a given point
		'
		Dim A, B As Double
		
		'Convert from degees to radians
		nAngle = nAngle * PI / 180
		
		B = System.Math.Sin(nAngle) * nLength
		A = System.Math.Cos(nAngle) * nLength
		
		xyReturn.X = xyStart.X + A
		xyReturn.y = xyStart.y + B
		
	End Sub
	
	Sub PR_CreateTapeLayer(ByRef sFileNo As String, ByRef sSide As String, ByRef nTape As Object)
		'To the DRAFIX macro file (given by the global fNum).
		'Write the syntax to create a layer for a sleeve tape label
		'This will be named by the following convention
		'    <"FileNo"> + <"L"|"R"> + <nTapeNo>
		'E.g. A1234567L10
		'
		'For this to work it assumes that the following DRAFIX variables
		'are defined.
		'    HANDLE  hLayer
		'
		'Note:-
		'    fNum, CC, QQ, NL, QCQ are globals initialised by FN_Open
		'
		Dim slayer As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nTape. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		slayer = sFileNo & Mid(sSide, 1, 1) & nTape
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hLayer = Table(" & QQ & "find" & QCQ & "layer" & QCQ & slayer & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "if ( hLayer != %badtable)" & "Execute (" & QQ & "menu" & QCQ & "SetLayer" & QC & "hLayer);")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "else")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Table (" & QQ & "add" & QCQ & "layer" & QCQ & slayer & QCQ & "Tape Layer Data" & QCQ & "current" & QC & "Table(" & QQ & "find" & QCQ & "color" & QCQ & "DarkCyan" & QQ & "));")
		
	End Sub
	
	Sub PR_DeleteText(ByRef sFileNo As Object, ByRef sSleeve As Object, ByRef nTape As Object)
		'For this to work it assumes that the following DRAFIX variables
		'are defined
		'    XY      xyStart
		'
		'Note:-
		'    fNum, CC, QQ, NL, g_nCurrTextAspect are globals initialised by FN_Open
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "nTape = " & nTape & ";")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "sFileNo = " & QQ & sFileNo & QQ & ";")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "sSleeve = " & QQ & sSleeve & QQ & ";")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "sLayer=" & QQ & "layer = '" & QQ & "+ sFileNo + StringMiddle(sSleeve, 1, 1) + MakeString(" & QQ & "long" & QQ & ", nTape) + " & QQ & "'" & QQ & ";")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hChan = Open(" & QQ & "selection" & QC & "sLayer);")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "UserSelection(" & QQ & "clear" & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "if (hChan){")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "ResetSelection(hChan);")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "while(hEnt = GetNextSelection (hChan)) DeleteEntity(hEnt);")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "}")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Close(" & QQ & "selection" & QC & "hChan);")
	End Sub
	
	Sub PR_DrawArc(ByRef xyCen As XY, ByRef xyArcStart As XY, ByRef xyArcEnd As XY)
		
		Dim nEndAng, nRad, nStartAng, nDeltaAng As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nRad. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nRad = FN_CalcLength(xyCen, xyArcStart)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nStartAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nStartAng = FN_CalcAngle(xyCen, xyArcStart)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nEndAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nEndAng = FN_CalcAngle(xyCen, xyArcEnd)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nStartAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object nEndAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object nDeltaAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nDeltaAng = nEndAng - nStartAng
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nDeltaAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object nStartAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object nRad. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hEnt = AddEntity(" & QQ & "arc" & QC & "xyStart.x +" & Str(xyCen.X) & CC & "xyStart.y +" & Str(xyCen.y) & CC & Str(nRad) & CC & Str(nStartAng) & CC & Str(nDeltaAng) & ");")
		
	End Sub
	
	Sub PR_DrawAssignDrafixVariable(ByRef sName As String, ByRef nValue As Double)
		'To the DRAFIX macro file (given by the global fNum).
		'Write the syntax to do a variable assignment
		'
		'For this to work it assumes that the following DRAFIX variables
		'are defined
		'    XY      xyStart
		'
		'Note:-
		'    fNum, CC, QQ, NL, g_nCurrTextAspect are globals initialised by FN_Open
		'
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, sName & "=" & nValue & ";")
		
		
	End Sub
	
	Sub PR_DrawCircle(ByRef xyCen As XY, ByRef nRadius As Object)
		'To the DRAFIX macro file (given by the global fNum).
		'Write the syntax to draw a CIRCLE at the point given.
		'For this to work it assumes that the following DRAFIX variables
		'are defined
		'    XY      xyStart
		'    HANDLE  hEnt
		'
		'Note:-
		'    fNum, CC, QQ, NL are globals initialised by FN_Open
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hEnt = AddEntity(" & QQ & "circle" & QC & "xyStart.x+" & Str(xyCen.X) & CC & "xyStart.y+" & Str(xyCen.y) & CC & nRadius & ");")
	End Sub
	
	Sub PR_DrawCircularStump(ByRef xyStart As XY, ByRef nFiguredLength As Double)
		'To the DRAFIX macro file (given by the global fNum)
		'Write the syntax to draw a CIRCULAR STUMP
		'
		'For this to work it assumes that the following DRAFIX variables
		'are defined
		'    XY      xyStart
		'    HANDLE  hEnt
		'
		'Note:-
		'    g_sFileno, g_sPatient, g_sSide are globals initialised by FN_Open
		'
		
		Dim xyPt As XY
		Dim nTextHt, nSeamAllowance As Single
		Dim nRadius As Single
		
		nTextHt = 0.125
		nSeamAllowance = 0.125 ' 1/8th Inch
		nRadius = (nFiguredLength / 3.14)
		
		PR_SetLayer("Template" & g_sSide)
		
		PR_DrawCircle(xyStart, nRadius + nSeamAllowance) 'Add seam allowance
		PR_AddEntityID(g_sFileNo, g_sSide, "CirStump") 'sFileNo, g_sSide from FN_Open
		
		PR_SetLayer("Notes")
		PR_DrawCircle(xyStart, nRadius)
		PR_AddEntityID(g_sFileNo, g_sSide, "CirStumpSeam") 'sFileNo, g_sSide from FN_Open
		
		PR_SetTextData(2, 16, -1, -1, -1) 'Horiz center, Vertical center
		PR_DrawText(g_sPatient & "\n" & g_sWorkOrder & "\n" & g_sSide, xyStart, nTextHt)
		
	End Sub
	
	Sub PR_DrawFitted(ByRef Profile As curve)
		'To the DRAFIX macro file (given by the global fNum)
		'write the syntax to draw a FITTED curve through the points
		'given in Profile.
		'For this to work it assumes that the following DRAFIX variables
		'are defined
		'    XY      xyStart
		'    HANDLE  hEnt
		'
		'Note:-
		'    fNum, CC, QQ, NL are globals initialised by FN_Open
		'
		'
		Dim ii As Short
		
		'Draw the profile
		'    If there is no vertex or only one vertex then exit.
		'    For two vertex draw as a polyline (this degenerates to a single line).
		'    For three vertex draw as a polyline (as no fitted curve can be drawn
		'    by a macro).
		'
		Select Case Profile.n
			Case 0 To 1
				Exit Sub
			Case 3
				PR_DrawPoly(Profile)
				'Warn the user to smooth the curve
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "Display (" & QQ & "message" & QCQ & "OKquestion" & QCQ & "The Profile has been drawn as a POLYLINE\nEdit this line and make it OPEN FITTED,\n this will then be a smooth line" & QQ & ");")
			Case Else
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, "hEnt = AddEntity(" & QQ & "poly" & QCQ & "fitted" & QQ)
				For ii = 1 To Profile.n
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, CC & "xyStart.x+" & Str(Profile.X(ii)) & CC & "xyStart.y+" & Str(Profile.y(ii)))
				Next ii
				'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(fNum, ");")
		End Select
		
	End Sub
	
	Sub PR_DrawLine(ByRef xyStart As XY, ByRef xyFinish As XY)
		'To the DRAFIX macro file (given by the global fNum).
		'Write the syntax to draw a LINE between two points.
		'For this to work it assumes that the following DRAFIX variables
		'are defined
		'    XY      xyStart
		'    HANDLE  hEnt
		'
		'Note:-
		'    fNum, CC, QQ, NL are globals initialised by FN_Open
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hEnt = AddEntity(")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, QQ & "line" & QC)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "xyStart.x+" & Str(xyStart.X) & CC & "xyStart.y+" & Str(xyStart.y) & CC)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "xyStart.x+" & Str(xyFinish.X) & CC & "xyStart.y+" & Str(xyFinish.y))
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, ");")
		
		
	End Sub
	
	Sub PR_DrawLineOffset(ByRef xyStart As XY, ByRef xyFinish As XY, ByRef nOffset As Double)
		'To the DRAFIX macro file (given by the global fNum)
		'write the syntax to draw a line between two points offset a given distance.
		'For this to work it assumes that xyStart is defined in DRAFIX
		'as an XY
		'
		'Note:-
		'    fNum, CC, QQ, NL are globals initialised by FN_Open.
		'    The line is always offset to the LEFT of the direction from start
		'    to finish.
		'    If nOffset is -ve then the line is offset to the RIGHT
		
		Dim nAngle As Double
		Dim xyPt1, xyPt2 As XY
		
		nAngle = FN_CalcAngle(xyStart, xyFinish)
		If nOffset < 0 Then
			nAngle = nAngle - 90
		Else
			nAngle = nAngle + 90
		End If
		
		PR_CalcPolar(xyStart, nAngle, System.Math.Abs(nOffset), xyPt1)
		PR_CalcPolar(xyFinish, nAngle, System.Math.Abs(nOffset), xyPt2)
		
		PR_DrawLine(xyPt1, xyPt2)
		
	End Sub
	
	Sub PR_DrawMesh()
		'Invokes external Mesh drawing routine
		'As such it must be the last line called by the DRAW Macro
		If g_bDrawBodyMesh Then
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "Execute (" & QQ & "application" & QC & "sPathJOBST + " & QQ & "\\raglan\\meshdraw" & QCQ & "normal" & QQ & " );")
		End If
		If g_bDrawVestMesh Then
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "Execute (" & QQ & "application" & QC & "sPathJOBST + " & QQ & "\\raglan\\meshvest" & QCQ & "normal" & QQ & " );")
		End If
	End Sub
	
	Sub PR_DrawPoly(ByRef Profile As curve)
		'To the DRAFIX macro file (given by the global fNum)
		'write the syntax to draw a POLYLINE through the points
		'given in Profile.
		'For this to work it assumes that the following DRAFIX variables
		'are defined
		'    XY      xyStart
		'    HANDLE  hEnt
		'
		'Note:-
		'    fNum, CC, QQ, NL are globals initialised by FN_Open
		'
		'
		Dim ii As Short
		
		'Exit if nothing to draw
		If Profile.n <= 1 Then Exit Sub
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hEnt = AddEntity(" & QQ & "poly" & QCQ & "polyline" & QQ)
		For ii = 1 To Profile.n
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, CC & "xyStart.x+" & Str(Profile.X(ii)) & CC & "xyStart.y+" & Str(Profile.y(ii)))
		Next ii
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, ");")
		
	End Sub
	
	Sub PR_DrawRaglan(ByRef sType As String, ByRef sVestRaglan As Object, ByRef sAge As Object, ByRef sVestID As Object)
		'To the DRAFIX macro file (given by the global fNum).
		'Write the syntax to add a Raglan to the end of the sleeve
		'profile based on the data given in the variable sVestRaglan
		'
		'Note:-
		' This routine is used to draw the raglan part of the
		' sleeve based on the data from the vest.
		' It is called VestRaglan to distinguish it from the
		' raglan flap.
		'
		'Input:-
		' sVestRaglan  A composite string consisting of the following:
		'
		'              sAxillaType, either Regular 2"
		'                                  Regular 1½"
		'                                  Regular 2½"
		'                                  Open
		'                                  Mesh
		'                                  Lining
		'                                  Sleeveless
		'              nAxillaFrontNeckRad
		'              nAxillaBackNeckRad
		'              nShoulderToBackRaglan
		'
		'              E.G. "Regular 1½" 5.565656 6.123344 0.3456
		'
		'Output:-
		' Calls to 3 different macro files written to fNum
		'    Open macro file.
		'
		'        nAxillaFrontNeckRad   }
		'        nAxillaBackNeckRad    } Data from sVestRaglan
		'        nShoulderToBackRaglan } (For vest only)
		'        sAxillaType           }
		'
		'    Axilla macro file (depending on axilla given).
		'    Closing macro file.
		'
		'
		'Note: This routine has been amended to support Bodysuits
		'      For vest read Vest & BodySuit
		Dim sAxillaType, sString As String
		Dim iPos As Short
		
		'Flag to draw the mesh used by PR_DrawMesh
		g_bDrawBodyMesh = False
		g_bDrawVestMesh = False
		
		
		'Find axilla type
		'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object sVestRaglan. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		iPos = InStr(sVestRaglan, QQ) 'Regular axilla are awkward as they
		'contain a space, look for quote (")
		'QQ is a constant set to quote (")
		If iPos <> 0 Then
			'Escape quote for use in drafix
			'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object sVestRaglan. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sAxillaType = Left(sVestRaglan, iPos - 1) & "\" & QQ
			'UPGRADE_WARNING: Couldn't resolve default property of object sVestRaglan. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sString = Mid(sVestRaglan, iPos + 1)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object sVestRaglan. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			iPos = InStr(sVestRaglan, " ")
			'UPGRADE_WARNING: Couldn't resolve default property of object sVestRaglan. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sAxillaType = Left(sVestRaglan, iPos - 1)
			'UPGRADE_WARNING: Couldn't resolve default property of object sVestRaglan. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sString = Mid(sVestRaglan, iPos)
		End If
		
		'Check that an axilla has been given
		If sAxillaType = "None" Or sAxillaType = "Sleeveless" Then Exit Sub
		
		'NB the order of the following is very important
		'
		'Data from vest used in drawing the sleeve
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "nAxillaFrontNeckRad= " & FN_GetNumber(sString, 1) & ";")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "nAxillaBackNeckRad= " & FN_GetNumber(sString, 2) & ";")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "nShoulderToBackRaglan = " & FN_GetNumber(sString, 3) & ";")
		
		If sType = "Vest" Then
			'Initialise raglan drawing
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\VR_INIT.D;")
			
			'Draw for the particular axilla type
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "sAxillaType = " & QQ & sAxillaType & QQ & ";")
			
			Select Case sAxillaType
				Case "Open", "Lining"
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\VR_OPEN.D;")
				Case "Mesh"
					g_bDrawVestMesh = True
					'UPGRADE_WARNING: Couldn't resolve default property of object sAge. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, "nAge = " & Val(sAge) & ";")
					If g_sSide = "Left" Then
						'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						PrintLine(fNum, "nMeshLength = " & Val(fnGetString(CType(armdia.Controls("txtMeshData"), Object).Text, 2, ",")) & ";")
						'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						PrintLine(fNum, "nDistanceAlongRaglan = " & Val(fnGetString(CType(armdia.Controls("txtMeshData"), Object).Text, 1, ",")) & ";")
					Else
						If Val(fnGetString(CType(armdia.Controls("txtMeshData"), Object).Text, 4, ",")) > 0 And Val(fnGetString(CType(armdia.Controls("txtMeshData"), Object).Text, 3, ",")) > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							PrintLine(fNum, "nMeshLength = " & Val(fnGetString(CType(armdia.Controls("txtMeshData"), Object).Text, 4, ",")) & ";")
							'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							PrintLine(fNum, "nDistanceAlongRaglan = " & Val(fnGetString(CType(armdia.Controls("txtMeshData"), Object).Text, 3, ",")) & ";")
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							PrintLine(fNum, "nMeshLength = " & Val(fnGetString(CType(armdia.Controls("txtMeshData"), Object).Text, 2, ",")) & ";")
							'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							PrintLine(fNum, "nDistanceAlongRaglan = " & Val(fnGetString(CType(armdia.Controls("txtMeshData"), Object).Text, 1, ",")) & ";")
						End If
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\VR_MESH.D;")
				Case Else 'Regular
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\VR_REGLR.D;")
			End Select
			
			'Close raglan drawing
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\VR_CLOSE.D;")
		Else
			'Initialise raglan drawing
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\BR_INIT.D;")
			
			'Draw for the particular axilla type
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "sAxillaType = " & QQ & sAxillaType & QQ & ";")
			
			Select Case sAxillaType
				Case "Open", "Lining"
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\BR_OPEN.D;")
				Case "Mesh"
					g_bDrawBodyMesh = True
					'UPGRADE_WARNING: Couldn't resolve default property of object sAge. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, "nAge = " & Val(sAge) & ";")
					If g_sSide = "Left" Then
						'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						PrintLine(fNum, "nMeshLength = " & Val(fnGetString(CType(armdia.Controls("txtMeshData"), Object).Text, 2, ",")) & ";")
						'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						PrintLine(fNum, "nDistanceAlongRaglan = " & Val(fnGetString(CType(armdia.Controls("txtMeshData"), Object).Text, 1, ",")) & ";")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						PrintLine(fNum, "nMeshLength = " & Val(fnGetString(CType(armdia.Controls("txtMeshData"), Object).Text, 4, ",")) & ";")
						'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						PrintLine(fNum, "nDistanceAlongRaglan = " & Val(fnGetString(CType(armdia.Controls("txtMeshData"), Object).Text, 3, ",")) & ";")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\BR_MESH.D;")
				Case Else 'Regular
					'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\BR_REGLR.D;")
			End Select
			
			'Close raglan drawing
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "@" & g_sPathJOBST & "\RAGLAN\BR_CLOSE.D;")
			
		End If
	End Sub
	
	Sub PR_DrawText(ByRef sText As Object, ByRef xyInsert As XY, ByRef nHeight As Object)
		'To the DRAFIX macro file (given by the global fNum).
		'Write the syntax to draw TEXT at the given height.
		'
		'For this to work it assumes that the following DRAFIX variables
		'are defined
		'    XY      xyStart
		'
		'Note:-
		'    fNum, CC, QQ, NL, g_nCurrTextAspect are globals initialised by FN_Open
		'
		'
		Dim nWidth As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object nHeight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object nWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nWidth = nHeight * g_nCurrTextAspect
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "AddEntity(" & QQ & "text" & QCQ & sText & QC & "xyStart.x+" & Str(xyInsert.X) & CC & "xyStart.y+" & Str(xyInsert.y) & CC & nWidth & CC & nHeight & CC & g_nCurrTextAngle & ");")
		
	End Sub
	
	Sub PR_GetStyle(ByRef sStyle As String)
		'This function returns a value of the following
		'in sWO
		'WorkOrder, ItemNo
		'eg
		'    12345,10
		'
		Dim nOptions As Short
		Dim sDistalStyle, sProximalStyle As String
		Dim sBestCatNo, sStyleWO As String
		Dim nResult As Short
		
		'Distal
		sDistalStyle = "XX"
		If g_iStyleFirstTape = g_iFirstTape Then
			'Plain
			sDistalStyle = "PL"
		Else
			'Start at tape
			If g_iStyleFirstTape < 10 Then
				sDistalStyle = "0" & LTrim(Str(g_iStyleFirstTape))
			Else
				sDistalStyle = LTrim(Str(g_iStyleFirstTape))
			End If
		End If
		If g_sGaunt = "True" Then
			'Gauntlet
			sDistalStyle = "GT"
		End If
		If g_sStump = "True" Then
			'Stump
			sDistalStyle = "ST"
		End If
		
		'Proximal
		sProximalStyle = "XX"
		If g_iStyleLastTape = g_iLastTape Then
			'Plain
			sProximalStyle = "PL"
		Else
			'Start at tape
			If g_iStyleLastTape < 10 Then
				sProximalStyle = "0" & LTrim(Str(g_iStyleLastTape))
			Else
				sProximalStyle = LTrim(Str(g_iStyleLastTape))
			End If
		End If
		If g_sDetGaunt = "True" And g_iStyleLastTape <= 6 Then
			'Detachable Gauntlet only, if drawn from or less than tape +1-1/2
			sProximalStyle = "GT"
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object g_sFlapChk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If g_sFlapChk = "True" Then
			'Flap
			sProximalStyle = "FP"
		End If
		If Left(armdia.cboFlaps.Text, 4) = "Vest" Then
			'Vest raglan
			sProximalStyle = "VR"
		End If
		If Left(armdia.cboFlaps.Text, 4) = "Body" Then
			'Bodysuit raglan
			sProximalStyle = "BR"
		End If
		
		'Join distal and proximal styles
		sStyle = sDistalStyle & sProximalStyle
		
		
	End Sub
	
	Sub PR_InsertSymbol(ByRef sSymbol As String, ByRef xyInsert As XY, ByRef nScale As Single, ByRef nRotation As Single, ByRef slayer As String)
		'To the DRAFIX macro file (given by the global fNum).
		'Write the syntax to insert a SYMBOL.
		'Where:-
		'    sSymbol     Symbol name, must exist and be in the symbol library
		'    xyInsert    The insertion point
		'    nScale      Symbol scaling factor, 1 = No scaling
		'    nRotation   Symbol rotation about insertion point
		'For this to work it assumes that the following DRAFIX variables
		'are defined
		'    XY      xyStart
		'    HANDLE  hEnt
		'and
		'    The DRAFIX symbol library "C:\JOBST\JOBST.SLB" exists
		'
		'Note:-
		'    fNum, CC, QQ, NL, QCQ are globals initialised by FN_Open
		'
		PR_SetLayer("Notes")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "SetSymbolLibrary(sPathJOBST +" & QQ & "\\JOBST.SLB" & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "Symbol(" & QQ & "find" & QCQ & sSymbol & QQ & ");")
		
		'PR_SetLayer slayer
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hEnt = AddEntity(" & QQ & "symbol" & QCQ & sSymbol & QC)
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "xyStart.x+" & Str(xyInsert.X) & CC & "xyStart.y+" & Str(xyInsert.y) & CC)
		'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, Str(nScale) & CC & Str(nScale) & CC & Str(nRotation) & ");")
		
	End Sub
	
	Sub PR_MakeXY(ByRef xyReturn As XY, ByRef X As Double, ByRef y As Double)
		'Utility to return a point based on the X and Y values
		'given
		xyReturn.X = X
		xyReturn.y = y
	End Sub
	
	Sub PR_NamedHandle(ByRef sHandleName As String)
		'To the DRAFIX macro file (given by the global fNum).
		'Write the syntax to retain the entity handle of a previously
		'added entity.
		'
		'Assumes that hEnt is the entity handle to the just inserted entity.
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "HANDLE " & sHandleName & ";")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, sHandleName & " = hEnt;")
		
	End Sub
	
	Sub PR_PutTapeLabel(ByRef nTape As Object, ByRef xyStart As XY, ByRef nLength As Object, ByRef nMM As Object, ByRef nGrm As Object, ByRef nRed As Object)
		Dim nDec As Object
		Dim nInt As Object
		'To the DRAFIX macro file (given by the global fNum).
		'Write the syntax to add Sleeve Tape details,
		'these details are given explicitly as arguments.
		'Where:-
		'    nTape       Index into sTextList below
		'    xyStart     Position of tape label on fold
		'    nLength     Tape length to be displayed, decimal inches
		'    nMM         MMs to be displayed
		'    nGrm        Grams to be displayed
		'    nRed        Reduction to be displayed
		'
		'
		'For this to work it assumes that the following DRAFIX variables
		'are defined
		'    XY      xyStart
		'    HANDLE  hEnt
		'
		'Note:-
		'    fNum, g_sFileNo, g_sSide are globals initialised by FN_Open
		'
		'
		Dim sTextList, sSymbol As String
		Dim xyPt As XY
		Dim nSymbolOffSet, nTextHt As Single
		
		'sTextList = "  6 4½  3 1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nTape. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sSymbol = nTape & "tape"
		nSymbolOffSet = 0.6877
		nTextHt = 0.125
		
		PR_MakeXY(xyPt, xyStart.X, xyStart.y + nSymbolOffSet) 'Offset because symbol point is at top
		PR_InsertSymbol(sSymbol, xyPt, 1, 0, "Notes")
		
		PR_CreateTapeLayer(g_sFileNo, g_sSide, nTape)
		
		PR_SetTextData(1, 32, -1, -1, 0)
		
		'Length text
		'N.B. format as Inches and eighths. With eighths offset up and left
		'UPGRADE_WARNING: Couldn't resolve default property of object nInt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nInt = Int(CDbl(nLength)) 'Integer part of the length (before decimal point)
		
		'Decimal part of the length (after decimal point)
		'convert to 1/8ths and get nearest by rounding
		'UPGRADE_WARNING: Couldn't resolve default property of object nInt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object nLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object nDec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nDec = round((nLength - nInt) / 0.125)
		'UPGRADE_WARNING: Couldn't resolve default property of object nDec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If nDec = 8 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object nDec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			nDec = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object nInt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			nInt = nInt + 1
		End If
		
		'Draw Integer part
		PR_MakeXY(xyPt, xyStart.X + 0.0625, xyStart.y + 0.75)
		'UPGRADE_WARNING: Couldn't resolve default property of object nInt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PR_DrawText(Trim(nInt), xyPt, nTextHt)
		
		'Draw eights part
		'UPGRADE_WARNING: Couldn't resolve default property of object nInt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PR_MakeXY(xyPt, xyStart.X + 0.0625 + (Len(Trim(nInt)) * nTextHt * 0.8), xyStart.y + 0.75 + nTextHt / 1.5)
		'UPGRADE_WARNING: Couldn't resolve default property of object nDec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If nDec <> 0 Then PR_DrawText(Trim(nDec), xyPt, nTextHt / 1.5)
		
		'MMs text
		PR_MakeXY(xyPt, xyStart.X + 0.0625, xyStart.y + 1)
		'UPGRADE_WARNING: Couldn't resolve default property of object nMM. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PR_DrawText(nMM & "mm", xyPt, nTextHt)
		
		'Grams text
		PR_MakeXY(xyPt, xyStart.X + 0.0625, xyStart.y + 1.25)
		'UPGRADE_WARNING: Couldn't resolve default property of object nGrm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PR_DrawText(nGrm & "gm", xyPt, nTextHt)
		
		'Reduction text and circle round the text
		PR_SetTextData(2, 16, -1, -1, -1)
		PR_MakeXY(xyPt, xyStart.X + 0.25, xyStart.y + 1.625)
		'UPGRADE_WARNING: Couldn't resolve default property of object nRed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PR_DrawText(Trim(nRed), xyPt, nTextHt)
		PR_DrawCircle(xyPt, 0.125)
		
		
		
	End Sub
	
	Sub PR_RightThumbHole(ByRef xyCen As XY, ByRef xyArcStart As XY, ByRef xyArcEnd As XY)
		' This Calculates and Draws the curve at the
		' right of the ThumbHole on the template
		
		Dim nEndAng, nRad, nStartAng, nDeltaAng As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object nRad. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nRad = FN_CalcLength(xyCen, xyArcStart)
		'UPGRADE_WARNING: Couldn't resolve default property of object nStartAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nStartAng = FN_CalcAngle(xyCen, xyArcStart)
		'UPGRADE_WARNING: Couldn't resolve default property of object nEndAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nEndAng = FN_CalcAngle(xyCen, xyArcEnd)
		'UPGRADE_WARNING: Couldn't resolve default property of object nStartAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object nEndAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object nDeltaAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nDeltaAng = nEndAng - nStartAng
		'UPGRADE_WARNING: Couldn't resolve default property of object nDeltaAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		nDeltaAng = -nDeltaAng
		'UPGRADE_WARNING: Couldn't resolve default property of object nDeltaAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object nStartAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object nRad. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hEnt = AddEntity(" & QQ & "arc" & QC & "xyStart.x +" & Str(xyCen.X) & CC & "xyStart.y +" & Str(xyCen.y) & CC & Str(nRad) & CC & Str(nStartAng) & CC & Str(nDeltaAng) & ");")
	End Sub
	
	Sub PR_SetLayer(ByRef sNewLayer As String)
		'To the DRAFIX macro file (given by the global fNum).
		'Write the syntax to set the current LAYER.
		'For this to work it assumes that hLayer is defined in DRAFIX as
		'a HANDLE.
		'
		'Note:-
		'    fNum, CC, QQ, NL, g_sCurrentLayer are globals initialised by FN_Open
		'
		'To reduce unessesary writing of DRAFIX code check that the new layer
		'is different from the Current layer, change only if it is different.
		'
		
		' If g_sCurrentLayer = sNewLayer Then Exit Sub
		'g_sCurrentLayer = sNewLayer
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "hLayer = Table(" & QQ & "find" & QCQ & "layer" & QCQ & sNewLayer & QQ & ");")
		'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(fNum, "if ( hLayer != %badtable)" & "Execute (" & QQ & "menu" & QCQ & "SetLayer" & QC & "hLayer);")
		
	End Sub
	
	Sub PR_SetTextData(ByRef nHoriz As Object, ByRef nVert As Object, ByRef nHt As Object, ByRef nAspect As Object, ByRef nFont As Object)
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
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetData(" & QQ & "TextHorzJust" & QC & nHoriz & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object nHoriz. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextHorizJust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextHorizJust = nHoriz
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nVert. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextVertJust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If nVert >= 0 And g_nCurrTextVertJust <> nVert Then
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetData(" & QQ & "TextVertJust" & QC & nVert & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object nVert. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextVertJust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextVertJust = nVert
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nHt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextHt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If nHt >= 0 And g_nCurrTextHt <> nHt Then
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetData(" & QQ & "TextHeight" & QC & nHt & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object nHt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextHt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextHt = nHt
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If nAspect >= 0 And g_nCurrTextAspect <> nAspect Then
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetData(" & QQ & "TextAspect" & QC & nAspect & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object nAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextAspect = nAspect
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object nFont. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextFont. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If nFont >= 0 And g_nCurrTextFont <> nFont Then
			'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(fNum, "SetData(" & QQ & "TextFont" & QC & nFont & ");")
			'UPGRADE_WARNING: Couldn't resolve default property of object nFont. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextFont. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_nCurrTextFont = nFont
		End If
		
		
	End Sub
	
	Function round(ByVal nNumber As Single) As Short
		'Fuction to return the rounded value of a decimal number
		'E.G.
		'    round(1.35) = 1
		'    round(1.55) = 2
		'    round(2.50) = 3
		'    round(-2.50) = -3
		'
		
		Dim nInt, nSign As Short
		
		nSign = System.Math.Sign(nNumber)
		nNumber = System.Math.Abs(nNumber)
		nInt = Int(nNumber)
		If (nNumber - nInt) >= 0.5 Then
			round = (nInt + 1) * nSign
		Else
			round = nInt * nSign
		End If
		
	End Function
End Module
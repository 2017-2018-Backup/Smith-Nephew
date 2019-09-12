// File Name:	WHFTREFC.D
// Purpose:	Draws heel reinforcement
// Version:	2.01
// Date:		17.Aug.1994
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
// 22.Sept.94	GG	Production evaluation modifications
// 06.0ct.94	GG	Made standalone & interactive to reduce code in WH_LEG1 & WH_LEG2
// 04.Jan.94	GG	Added "All Around" and "Through End"
// 21.Nov.96	GG   	No longer looks for leg box 

HANDLE
	hDlg,
	hEnt,
	hChan,
	hHeelMarker,
	hLayer;

XY
	xyInt,	// Returned by FN_CirLinInt
	xyHeel,
	xyOtemplate,
	xyAnkle,
	xyAnkleM,
	xyPt1,
	xyTmp;

STRING
	sHeelStyle,
	sProfileID,
	sHeelID,
	sOtemplateID,
	sAnkleID,
	sAnkleMID,
	sTmp,
	sHeelLabel,
	sLeg,
	sFileNo,
	sStyleList,
	sStyleID,
	sDlgStyleList,
	sSymbolBox,
	sType;

NUMBER
	nAnkleX, nAnkleY,
	nAnkleMX, nAnkleMY,
	nLowLine,
	nHighLine;

BOOL
	SmallHeel;

MSGID
	Ok;

ANGLE
	aAngle;
	
// Find JOBST installed directory
// Set path to macros
//
   STRING	 sPathJOBST;
   sPathJOBST = GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;
   SetData("PathMacro",  sPathJOBST + "\\WAIST" );
   
   UserSelection("clear") ;
   UserSelection("update") ;
   Execute ( "menu", "SetStyle", Table( "find",  "style", "bylayer"));
   Execute ( "menu", "SetColor", Table( "find",  "color", "bylayer"));


// Sub Routines etc
//
   @WHFTRDEF.D;

// Get profile, identify Leg and FileNo

   hEnt = GetEntity ("Select a Leg Profile", "curve") ;

   Display ("cursor", "wait", "Searching for more data");
   GetDBData (hEnt, "ID", &sProfileID);
   nStringLength = StringLength(sProfileID);
   if (StringCompare( "LeftLegCurve", StringMiddle(sProfileID, nStringLength - 11, 12))){
	sStyleID = StringMiddle(sProfileID, 1, nStringLength - 12);
	sLeg = "Left";
	}
    if (StringCompare( "RightLegCurve", StringMiddle(sProfileID, nStringLength - 12, 13))) {
	sStyleID = StringMiddle(sProfileID, 1, nStringLength - 13);
	sLeg = "Right";
	}

// Check if this is a valid Leg Curve exit if not
    if (StringLength (sLeg) == 0) Exit (%cancel, "A Leg Profile was not selected\n") ;

// Get Marker data
   sHeelID          = sStyleID + sLeg+ "Heel";
   sOtemplateID = sStyleID + sLeg + "Origin";
   sAnkleID        =  sStyleID + sLeg + "Ankle";
   sAnkleMID     =  sStyleID + sLeg + "AnkleM";

   sTmp = "type = 'Marker' AND (DB ID ='" + sHeelID + "' OR DB ID ='"  + sAnkleID + "' OR DB ID ='"  + sAnkleMID + "' OR DB ID ='" + sOtemplateID + "')";

   nMarkersFound = 0 ;
   nAllAroundMarkersFound = 0 ; // Special markers for "All Around"
   hChan = Open ("selection", sTmp );
   if (hChan) {
	ResetSelection(hChan);
	while (hEnt = GetNextSelection(hChan)) {
		GetDBData(hEnt,"ID", &sTmp) ;
		GetGeometry(hEnt, &sType, &xyPt1, &xyTmp, &aAngle) ;
	                if ( xyPt1.x != 0 || xyPt1.y != 0) {
			if (StringCompare(sTmp, sHeelID )){
				nMarkersFound = nMarkersFound + 2 ;
				GetGeometry(hEnt, &sTmp, &xyHeel, &xyTmp, &aAngle) ;
				GetDBData(hEnt, "Data", &sTmp);
   				GetDBData(hEnt, "HeelStyle", &sHeelStyle); 	
   				hHeelMarker = hEnt ;
				SmallHeel = Value("scalar", sTmp);
				}
			if (StringCompare(sTmp, sOtemplateID)) {
				nMarkersFound = nMarkersFound + 1 ;
				GetGeometry(hEnt, &sTmp, &xyOtemplate , &xyTmp, &aAngle) ;
				}
			if (StringCompare(sTmp, sAnkleID)) {
				 // Special markers for "All Around"
				nAllAroundMarkersFound = nAllAroundMarkersFound + 1 ;
				GetGeometry(hEnt, &sTmp, &xyAnkle , &xyTmp, &aAngle) ;
				GetDBData(hEnt, "Data", &sTmp);
				if ( !ScanLine(sTmp, "blank", &nAnkleX, &nAnkleY)) Exit(%cancel,"Can't get data from Ankle Marker!\n");
				}
			if (StringCompare(sTmp, sAnkleMID)) {
 				// Special markers for "All Around"
				nAllAroundMarkersFound = nAllAroundMarkersFound + 1 ;
				GetGeometry(hEnt, &sTmp, &xyAnkleM , &xyTmp, &aAngle) ;
				GetDBData(hEnt, "Data", &sTmp);
				if ( !ScanLine(sTmp, "blank", &nAnkleMX, &nAnkleMY)) Exit(%cancel,"Can't get data from Distal Ankle Marker!\n");
				}
			}
		}
   	}
   Close("selection", hChan); 

// Check if that the markers have been found, otherwise exit
//
    if (nMarkersFound == 0) Exit (%cancel, "Marker data not found!\n") ;
    if (nMarkersFound == 1) Exit (%cancel, "Heel Marker not found!\n") ;
    if (nMarkersFound == 2) Exit (%cancel, "Template Origin Marker not Found! \n") ;

// Create Dialog
// Get heel style
//
   sDlgStyleList = sHeelStyle + "\nRegular\nExtra Heel\nExtend on Foot\nBoth Extra & Extend\nAll Around\nThrough End";
   nButX = 10; nButY = 40;
   hDlg = Open ("dialog", sLeg + " Heel Reinforcement", "font Helv 8", 20,20, 115, 70);

	AddControl ( hDlg, "pushbutton",  nButX, nButY, 40, 14, "Cancel", "%cancel", "");
	AddControl ( hDlg, "pushbutton",  nButX + 53 , nButY, 40,14, "Draw", "%ok", "");
	AddControl ( hDlg, "dropdownlist", 10, 10, 90, 50 , sDlgStyleList, "string", "sHeelStyle");

   Ok = Display("dialog", hDlg, "%center");
   Close("dialog", hDlg);

// Draw Chosen foot reinforcement style
//
   if (Ok != %cancel && StringLength (sHeelStyle) != 0) {
    	hLayer = Table("find", "layer", "Notes") ;
    	if ( hLayer > %zero )
 		Execute ("menu", "SetLayer",  hLayer) ;
   	 else
   		Table ("add", "layer", sLayer, "Notes in Green", "current", Table("find", "color", "Green"));

         	if (SmallHeel) {
		nLowLine = 1.125;	// Default to regular heel
		nHighLine = 1.125;
		nHeelR1 = 1.1760709 ;  // See also WHFTPNTS.D
		nHeelR3 = 1.1760709 ; 
		nTextPos = .5;
		if ( StringCompare ("Extra Heel", sHeelStyle) 
		     || StringCompare ("Both Extra & Extend", sHeelStyle)) nHighLine = 1.5;
		if ( StringCompare ("Extend on Foot", sHeelStyle) 
		     || StringCompare ("Both Extra & Extend", sHeelStyle)) nLowLine = 1.5;
		}
	else {
		nLowLine = 1.875;
		nHighLine = 1.875;
		nHeelR1 = 1.6601037 ;
		nHeelR3 = 1.5053069 ;
		nHeelOff = 0.5;
		nTextPos = 1;
		if ( StringCompare ("Extra Heel", sHeelStyle) 
		     || StringCompare ("Both Extra & Extend", sHeelStyle)) nHighLine = 2.71875;
		if ( StringCompare ("Extend on Foot", sHeelStyle) 
		     || StringCompare ("Both Extra & Extend", sHeelStyle)) nLowLine = 2.71875;
		}

	// Draw lines, As dotted line (except for through end)
	Execute ("menu", "SetStyle",  Table("find", "style", "Long Dash")) ;
	if (!StringCompare ("Through End", sHeelStyle))
		AddEntity("line", xyHeel.x - nLowLine, xyHeel.y, xyHeel.x - nLowLine, xyOtemplate.y);
	AddEntity("line", xyHeel.x + nHighLine, xyHeel.y, xyHeel.x + nHighLine, xyOtemplate.y);
	Execute ("menu", "SetStyle",  Table("find", "style", "bylayer")) ;

	sHeelLabel = "REINFORCED" ;
	if ( StringCompare ("Extra Heel", sHeelStyle)  || StringCompare ("Both Extra & Extend", sHeelStyle) ){
		xyTmp.x = xyHeel.x + nHighLine ;
		xyTmp.y = xyOtemplate.y + (xyHeel.y - xyOtemplate.y) / 2 ;
   		AddEntity("marker", "closed arrow", xyTmp, 0.5 , 0.125, 180 ) ;
		sHeelLabel = sHeelLabel + " TO ARROW" ;
		}

	SetData("TextHorzJust", 2) ;
	AddEntity("text", sHeelLabel , xyHeel.x, xyHeel.y - nTextPos);

	if (StringCompare ("Through End", sHeelStyle))
		AddEntity("text", "THROUGH END" , xyHeel.x,( xyHeel.y - nTextPos) - 1.25);

	if (StringCompare ("All Around", sHeelStyle)){

		AddEntity("text", "ALL AROUND" , xyHeel.x, (xyHeel.y - nTextPos) - 1.25);

	   	AddEntity("marker", "closed arrow", xyHeel.x - nLowLine, xyOtemplate.y, 0.5 , 0.125, 45 ) ;		   
	   	AddEntity("marker", "closed arrow",xyHeel.x + nHighLine, xyOtemplate.y, 0.5 , 0.125, 135 ) ;		   

		if (nAllAroundMarkersFound != 2)
			Exit(%cancel,"Can't get data from Ankle Markers!\nCan't finish ALL AROUND reinforcement.");
		nError = FN_CirLinInt( xyHeel.x - nLowLine, xyHeel.y, 
				    xyHeel.x - nLowLine, xyOtemplate.y,
				    xyAnkleM.x + nAnkleMX, xyAnkleM.y + nAnkleMY, 
				    nHeelR1) ;
	   	AddEntity("marker", "closed arrow", xyInt, 0.5 , 0.125, 315 ) ;
		nError = FN_CirLinInt( xyHeel.x + nHighLine, xyHeel.y,
				    xyHeel.x + nHighLine, xyOtemplate.y,
				    xyAnkle.x + nAnkleX, xyAnkle.y + nAnkleY, 
				    nHeelR3) ;
	   	AddEntity("marker", "closed arrow", xyInt, 0.5 , 0.125, 225 ) ;		   
		}

	// Update Heel marker 
	SetDBData (hHeelMarker , "HeelStyle", sHeelStyle); 
    }

// Reset and exit
   Execute ("menu", "SetLayer",  Table("find", "layer", "1")) ;
   Exit (%ok, "Reinforcement drawing Complete");

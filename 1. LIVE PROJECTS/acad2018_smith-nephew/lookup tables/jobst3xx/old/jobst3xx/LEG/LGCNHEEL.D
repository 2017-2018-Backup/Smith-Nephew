// File Name:	LGCNHEEL.D
// Purpose:	Draws a HEEL contracture
// Version:	1.01
// Date:		06.Jan.95
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action

   Display ("cursor", "wait", "Loading");

HANDLE
	hCurv,
	hDlg,
	hEnt,
	hChan,
	hLayer;

XY
	xyInt,	// Returned by FN_CirLinInt
	xyOtemplate,
	xyHeel,
	xyPt1,
	xyPt2,
	xyPt3,
	xyTmp;

STRING
	sProfileID,
	sStyleID,
	sOtemplateID,
	sHeelID,
	sUnits,
	sTmp,
	sLeg,
	sFileNo,
	sType;

BOOL
	SmallHeel;
ANGLE
	aAngle;
NUMBER
	ii;
MSGID
	Ok;
	
// Find JOBST installed directory
// Set path to macros
//
   STRING	 sPathJOBST;
   sPathJOBST = GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;
   SetData("PathMacro",  sPathJOBST + "\\LEG" );
 
// Sub Routines etc
//
   @LGCNDEF.D;

// Reset 
    UserSelection ("clear");
    Execute ( "menu", "SetStyle", Table( "find",  "style", "bylayer"));
    Execute ( "menu", "SetColor", Table( "find",  "color", "bylayer"));

// Get profile, identify Leg and FileNo

   if(!hCurv = GetEntity ("Select a Leg Profile", "curve")) Exit (%ok, "User cancelled") ;

   Display ("cursor", "wait", "Searching for more data");
   GetDBData (hCurv, "ID", &sProfileID);
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
   nMarkersFound = 0 ;

   sOtemplateID = sStyleID + sLeg + "Origin";
   sHeelID = sStyleID + sLeg + "Heel";
   sTmp = "type = 'Marker' AND (DB ID ='" + sHeelID +  "' OR DB ID ='" + sOtemplateID + "')";
   hChan = Open ("selection", sTmp );
   if (hChan) {
	ResetSelection(hChan);
	while (hEnt = GetNextSelection(hChan)) {
		GetDBData(hEnt,"ID", &sTmp) ;
		GetGeometry(hEnt, &sType, &xyPt1, &xyTmp, &aAngle) ;
	                if ( xyPt1.x != 0 || xyPt1.y != 0) {
			if (StringCompare(sTmp, sOtemplateID)) {
				nMarkersFound = nMarkersFound + 1 ;
				GetGeometry(hEnt, &sTmp, &xyOtemplate , &xyTmp, &aAngle) ;
				}
			if (StringCompare(sTmp, sHeelID)) {
				nMarkersFound = nMarkersFound + 1 ;
				GetGeometry(hEnt, &sTmp, &xyHeel , &xyTmp, &aAngle) ;
				GetDBData(hEnt, "Data", &sTmp);
				SmallHeel = Value("scalar", sTmp);
				}
			}
		}
   	}
   Close("selection", hChan); 

// Check if the markers have been found, otherwise exit
//
    if (nMarkersFound < 2) Exit (%cancel, "Missing markers for selected foot, data not found!\n") ;
    if (nMarkersFound > 2) Exit (%cancel, "Two or more drawings of the same style exist!\nDelete the extra drawing/s and try again.\n") ;

// Calculate contracture points
   if (SmallHeel) 
	nHeelOff = 0.25 ;
   else
	nHeelOff = 0.5 ;

   nHeelCir = xyHeel.y - xyOtemplate.y ;

   xyPt1.x = xyHeel.x - nHeelOff ;
   xyPt1.y = xyOtemplate.y ;

   xyPt2.x = xyHeel.x ;
   xyPt2.y = xyOtemplate.y + nSeam + (nHeelCir /3) ;

   xyPt3.x = xyHeel.x + nHeelOff ;
   xyPt3.y = xyOtemplate.y ;
	
// Draw contracture
//
   PRSetLayerNotes();
   StartPoly("polyline");
     AddVertex (xyPt1);
     AddVertex (xyPt2);
     AddVertex (xyPt3);
  EndPoly();

   SetData("TextHorzJust", 1) ;
   AddEntity("text", "CONTRACTURE" , xyPt3.x, xyPt3.y + 0.5 );

// Reset and exit
   Execute ("menu", "SetLayer",  Table("find", "layer", "1")) ;
   Exit (%ok, "Contracture drawing Complete");

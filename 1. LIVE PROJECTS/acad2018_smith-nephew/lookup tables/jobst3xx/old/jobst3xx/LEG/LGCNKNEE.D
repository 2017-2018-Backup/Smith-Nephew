// File Name:	LGCNKNEE.D
// Purpose:	Draws a Knee contracture
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
	xyFirstProfile,
	xyPt1,
	xyPt2,
	xyKneeDip,
	xyTmp;

STRING
	sDlgContractureList,
	sContracture,
	sProfileID,
	sStyleID,
	sOtemplateID,
	sUnits,
	sTmp,
	sLeg,
	sType;

BOOL
	OriginFound;
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
   sOtemplateID = sStyleID + sLeg + "Origin";
   OriginFound = %false ;
   sTmp = "type = 'Marker' AND ( DB ID ='" + sOtemplateID + "')";

   hChan = Open ("selection", sTmp );
   if (hChan) {
	ResetSelection(hChan);
	while (hEnt = GetNextSelection(hChan)) {
		GetDBData(hEnt,"ID", &sTmp) ;
		GetGeometry(hEnt, &sType, &xyPt1, &xyTmp, &aAngle) ;
	                if ( xyPt1.x != 0 || xyPt1.y != 0) {
			if (StringCompare(sTmp, sOtemplateID)) {
				OriginFound = %true ;
				GetGeometry(hEnt, &sTmp, &xyOtemplate , &xyTmp, &aAngle) ;
				}
			}
		}
   	}
   Close("selection", hChan); 

// Check if that the markers have been found, otherwise exit
//
    if (!OriginFound) Exit (%cancel, "Origin Marker not found!\n") ;

// Start dialog wrt contractures
//
   sDlgContractureList =  "\n10 to 35 Degrees\n36 to 70 Degrees\n71 Degrees & Over ";
   nButX = 10; nButY = 40;
   hDlg = Open ("dialog", sLeg + " Knee Contractures", "font Helv 8", 20,20, 115, 70);

	AddControl ( hDlg, "pushbutton",  nButX, nButY, 40, 14, "Cancel", "%cancel", "");
	AddControl ( hDlg, "pushbutton",  nButX + 53 , nButY, 40,14, "Draw", "%ok", "");
	AddControl ( hDlg, "dropdownlist", 10, 10, 90, 50 , sDlgContractureList , "string", "sContracture");

   Ok = Display("dialog", hDlg, "%center");
   Close("dialog", hDlg);
   
   if (Ok == %cancel ) Exit (%ok, "User Cancel!") ;	

   sTmp = StringMiddle(sContracture,1, 2);
   if ( StringCompare ("10", sTmp)) nContracture = 0.5;
   if ( StringCompare ("36", sTmp)) nContracture = 1;
   if ( StringCompare ("71", sTmp)) nContracture = 1.5;

   hEnt = GetEntity ("Select the Knee Dip tape symbol", "symbol") ;
   GetGeometry(hEnt, &sType, &xyKneeDip, &xyTmp, &aAngle);

// Loop through profile and find the kneedip vertex
//
   nVertex =  GetVertexCount (hCurv);
   nn = 1;
   while (nn <= nVertex) {
	GetVertex ( hCurv, nn, &xyTmp) ;
	nTolerance = abs(xyTmp.x - xyKneeDip.x) ;
	if (nTolerance < 0.1 ){
		xyKneeDip = xyTmp;
 		break;
		}
	nn = nn + 1;
	}
   if (nn == nVertex + 1 ) Exit (%cancel, "Can't find Knee Dip vertex on selected profile.\n") ;
  
// Get next vertex to allow construction of contracture.
// Check availability of vertices to avoid nasty failure and error messages.
// Note:-
//       Special case where contracture width is longer than the distance between the
//       initially selected next vertex and the knee dip.

   if (nn + 1 > nVertex  ) Exit (%cancel, "Can't calculate Knee contracture on selected profile.\n") ;		
   GetVertex ( hCurv, nn + 1, &xyPt1) ;

   nError = FN_CirLinInt (xyKneeDip, xyPt1, xyKneeDip, nContracture) ;
   if (nError == %false) {
   	if (nn + 2 > nVertex  ) Exit (%cancel, "Can't calculate Knee contracture on selected profile.\n") ;		
   	GetVertex ( hCurv, nn + 2, &xyPt2) ;
	nError = FN_CirLinInt (xyPt1, xyPt2, xyKneeDip, nContracture) ;
	if (nError == %false) Exit (%cancel, "Can't calculate Knee contracture on selected profile.\n") ;		
	}
	
   xyPt2 = xyInt ;

   nKneeCir = xyKneeDip.y - xyOtemplate.y;
   aAngle = Calc ("angle", xyKneeDip, xyPt2);   
   nA = sqrt ( (0.5*nKneeCir)^2 - (0.5*nContracture)^2);
   xyPt1 = CalcXY ("relpolar", CalcXY ("relpolar", xyKneeDip, nContracture * 0.5, aAngle) , nA, aAngle + 270) ;  

// Draw contracture
//
   PRSetLayerNotes();
   StartPoly("polyline");
     AddVertex (xyKneeDip);
     AddVertex (xyPt1);
     AddVertex (xyPt2);
   EndPoly();

   SetData("TextHorzJust", 4) ;
   AddEntity("text", "CONTRACTURE" , xyKneeDip.x, xyKneeDip.y - ((xyKneeDip.y - xyPt1.y) / 2) );

// Display Contracture size on layer construct (for info only)
   PRSetLayerConstruct();
   SetData("TextHorzJust", 2) ;
   AddEntity("text", sContracture , xyPt1.x, xyPt1.y - 0.25);  

// Reset and exit
   Execute ("menu", "SetLayer",  Table("find", "layer", "1")) ;
   Exit (%ok, "Contracture drawing Complete");

// File Name:	ARMTHUMB.D
// Purpose: 	Tool to adjust the thumb hole position
//		And Pass Control To Visual Basic
//		For Arm Details
// Version:	1.1
// Date:		19.Apr.95
//		G. George
HANDLE	
	hSel,
	hEnt,		
	hCurv,
	hConstruct,	
	hEntityLayer,	
	hChan;
XY	
	xyO, 
	xyTmp, 
	xyProfilePt, 
	xyTranslate,
	xyPt1, 
	xyStart, 
	xyEnd;
STRING
	sFileNo,		
	sData,		
	sOtemplateID ,
	sTmp,		
	sProfileID,	
	sStyle,		
	sSleeve,
	sType,		
	sEntityType;
	
NUMBER	
	aAngle;
	
BOOL
	Loop;
	
// Clear all user selections
//	
   UserSelection("clear") ;
   UserSelection("update") ;
   Execute ( "menu", "SetStyle", Table( "find",  "style", "bylayer"));
   Execute ( "menu", "SetColor", Table( "find",  "color", "bylayer"));

// Get sleeve profile, identify ARM and FileNo
//
   hCurv = GetEntity ("Select a Arm Profile", "curve") ;
   if (hCurv == %zero)Exit(%ok,"No Arm Profile selected");

   Display ("cursor", "wait", "Searching for more data");
   GetDBData (hCurv, "ID", &sProfileID);
   nStringLength = StringLength(sProfileID);
   if (StringCompare( "LeftProfile", StringMiddle(sProfileID, nStringLength - 10, 11))){
	sStyle = StringMiddle(sProfileID, 1, nStringLength - 11);
	sSleeve = "Left";
	sFileNo =  StringMiddle(sProfileID, 5, nStringLength - 15);
	}
    if (StringCompare( "RightProfile",StringMiddle(sProfileID, nStringLength - 11, 12))) {
	sStyle = StringMiddle(sProfileID, 1, nStringLength - 12);
	sSleeve = "Right";
	sFileNo =  StringMiddle(sProfileID, 5, nStringLength - 16);
	}

// Check if this is a valid Arm Curve exit if not
//
    if (StringLength (sSleeve) == 0) Exit (%cancel, "An ARM Profile was not selected\n") ;

// Locate origin marker for selected style
//
   sOtemplateID = sStyle + sSleeve +  "originmark";
   sTmp = "type = 'Marker' AND DB ID = '" + sOtemplateID + "'";
   hChan = Open ("selection", sTmp );
   nMarkersFound = 0;
   if (hChan) {
	ResetSelection(hChan);
	while (hEnt = GetNextSelection(hChan)) {
		GetDBData(hEnt,"ID", &sTmp) ;
		GetGeometry(hEnt, &sType, &xyPt1, &xyTmp, &aAngle) ;
	                if ( xyPt1.x != 0 || xyPt1.y != 0) {
			if (StringCompare(sTmp, sOtemplateID)) {
				nMarkersFound = nMarkersFound + 1;
				xyO = xyPt1;
				}
			}
		}
   	}
   Close("selection", hChan); 

// Check if that the markers have been found, otherwise exit
//
    if (nMarkersFound < 1) Exit (%cancel, "Missing origin marker for selected ARM, data not found!\n") ;
    if (nMarkersFound > 1) Exit (%cancel, "Two or more drawings of the same style exist!\nDelete the extra drawing/s and try again.\n") ;

// Having got the Origin we now get from the user the position on the profile 
// that can be used to ahjust the thumb hole position
   Loop = %true ;
   while (Loop) {
  	if (!GetUser("xy", "Select point on template profile!", &xyProfilePt )) Exit(%ok,"Point not selected");
	if ( xyProfilePt.y < xyO.y){
		Display ("beep", "Selected point is below edge of template. Try again");
		Loop = %true;
		 }
	else 
		Loop = %false;
	}


// Find the components of the thumbhole
// 
   sTmp = "DB ID = '" + sStyle + sSleeve + "ThumbHole'";
   sTmp = sTmp +  "OR (DB ID = '" + sStyle + sSleeve + "ThumbHoleConstruct' AND type = 'marker')"; 
Print ("\n", sTmp) ;  
   hChan = Open ("selection", sTmp );
   if (hChan) {
   	// loop through selection looking for lines and the Y value of the lowest line
   	// We know there are only 2 lines nad the lowes line form the bottom of
   	// the thumb hole. 
	ResetSelection(hChan);
	nYThumbBott = 1000000 ;  // Impossible value
	while (hEnt = GetNextSelection(hChan)) {
		UserSelection ("add", hEnt) ;
		GetEntityClass(hEnt, &sEntityType);
		GetEntityData (hEnt, "layer", &hEntityLayer);
		// Thumb Hole elements
		if (StringCompare("line", sEntityType)) {
			// Upper and Lower lines
			GetGeometry(hEnt,  &xyStart, &xyEnd) ;
			nYThumbBott = min (nYThumbBott,  xyStart.y, xyEnd.y) ;
			}
		if (StringCompare("marker", sEntityType)) 
			// Thumb Hole construction marker
			hConstruct = hEnt ;
		}
	// Check that the found Y value of the bottom is sensible
	//
	if ( nYThumbBott == 1000000 ||  nYThumbBott < xyO.y || nYThumbBott > xyProfilePt.y ){
		UserSelection("clear") ;
  		UserSelection("update") ;
		Close("selection", hChan); 
		Exit (%abort, "Can't get original Thumb Hole position");
		}
		
	// Translate
	//
	ResetSelection(hChan);
	nSeam = 0.1875;
	nNewYThumbBott = xyO.y + (((xyProfilePt.y -  xyO.y) - nSeam)/2) + nSeam ;
	xyTranslate.y =  nNewYThumbBott - nYThumbBott ;
	xyTranslate.x = 0 ;
	Transform ( "move", "translate", xyTranslate) ;
	}
   Close("selection", hChan); 

// Adjust Construction line
// 
   if (hConstruct) {
   	GetGeometry (hConstruct, &sType, &xyPt1, &xyTmp, &aAngle) ;
   	xyPt1.y = xyProfilePt.y ;
   	SetGeometry (hConstruct, sType, xyPt1, xyTmp, aAngle) ;
   	}

   UserSelection("clear") ;
   UserSelection("update") ;

// File Name:	WHZIPBOD.D
// Purpose:	Draws a Waist Height body zipper
//		Waist Height
// Version:	1.01
// Date:		13.Mar.1995
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action

HANDLE
	hDlg,
	hEnt,
	hEOS,
	hChan,
	hLeg,
	hCurve,
	hLayer;

XY
	xyHorizStart,
	xyHorizEnd,
	xyMax, xyMin,
	xyCO_WaistBott, 
	xyOtemplate,
	xyPt1,
	xyTmp;

STRING
	sClass,
	sElasticProximal,
	sMedial,
	sZipperStyle,
	sUnits,
	sZipLength,
	sZipOffset,
	sProfileID,
	sOtemplateID,
	sCO_WaistBottID,
	sTmp, sTmp1,
	sLeg,
	sStyleID,
	sDlgElasticList ,
	sDlgLengthList,
	sType;

NUMBER
	nMedial,
	nMinZipLength,
	nThighFiguredCir, WaistActualCir,
	nElasticProximal,
	nElastic,	//sum of Proximal and Distal
	nType,
	ii, nn,
	nZipOff,
	nZipLength ;
BOOL
	EOSFound,
	EOStoSelectedPoint,
	Loop ;
MSGID
	Ok;
	
// Reset 
    UserSelection ("clear");
    UserSelection("update") ;
    Execute ( "menu", "SetStyle", Table( "find",  "style", "bylayer"));
    Execute ( "menu", "SetColor", Table( "find",  "color", "bylayer"));

// Get profile, identify Leg and FileNo
//
   if(!hCurve = GetEntity ("Select a Leg Profile", "curve")) Exit (%ok, "User cancelled") ;

   Display ("cursor", "wait", "Searching for more data");
   GetDBData (hCurve, "ID", &sProfileID);
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
// Revise style ID to find body bits
//   
   sStyleID = StringMiddle(sStyleID, 5, StringLength(sStyleID)-4) ;
   sCO_WaistBottID = sStyleID + sLeg + "CO_WaistBott";
   sTmp = "type = 'Marker' AND (DB ID ='" + sCO_WaistBottID + "' OR DB ID ='" + sOtemplateID + "')";
   nMarkersFound = 0 ;
   hChan = Open ("selection", sTmp );
   if (hChan) {
	ResetSelection(hChan);
	while (hEnt = GetNextSelection(hChan)) {
		GetDBData(hEnt,"ID", &sTmp) ;
		GetGeometry(hEnt, &sTmp1, &xyPt1, &xyTmp, &aAngle) ;
	                if ( xyPt1.x != 0 || xyPt1.y != 0) {
			if (StringCompare(sTmp, sOtemplateID)) {
				nMarkersFound = nMarkersFound + 1 ;
				GetGeometry(hEnt, &sTmp, &xyOtemplate , &xyTmp, &aAngle) ;
				GetDBData(hEnt, "units", &sUnits);
				GetDBData(hEnt, "Data", &sTmp) ;
	  			if ( !ScanLine(sTmp, "blank", &nType, &nn)) Exit(%cancel,"Can't get data from Origin Marker!\n");				
				sType = sTmp;
				}
			if (StringCompare(sTmp, sCO_WaistBottID )){
				// This is only for first drawn leg
				GetGeometry(hEnt, &sTmp, &xyCO_WaistBott, &xyTmp, &aAngle) ;
				}
			}
		}
   	}
   Close("selection", hChan); 

// Check if that the markers have been found, otherwise exit
//
  if (StringCompare("CHAP", sType))Exit (%cancel, "Can't use this zipper for a CHAP style!\n") ;
  if (nMarkersFound < 1) Exit (%cancel, "Missing marker for selected Waist Height, data not found!\n") ;
  if (nMarkersFound > 1) Exit (%cancel, "Two or more drawings of the same style exist!\nDelete the extra drawing/s and try again.\n") ;

// Create Dialog
// Get Zipper style
//
   sUnits = "Inches";

   EOStoSelectedPoint = %false ;
   sDlgElasticList =  "3/8\" Elastic\n3/4\" Elastic\n1½\" Elastic\nNo Elastic";
   sDlgElasticList =  "3/4\" Elastic\n" + sDlgElasticList;
   sDlgLengthList =  "Give a length\nSelected Point";
   nMedial = 0 ;
   Loop = %true ;
   nMinZipLength = 5 ;
   while (Loop) {		
            nButX = 65; nButY = 40;
            hDlg = Open ("dialog", sLeg + " Body Zipper (Waist Height)", "font Helv 8", 20,20, 210, 75);

	AddControl ( hDlg, "pushbutton",  nButX, nButY, 35, 14, "Cancel", "%cancel", "");
	AddControl ( hDlg, "pushbutton",  nButX + 48 , nButY, 35,14, "OK", "%ok", "");

	AddControl ( hDlg, "ltext", 5, 12, 28, 14 , "EOS to", "string", "");
	AddControl ( hDlg, "combobox", 35, 10, 70, 40 , sDlgLengthList , "string", "sZipLength");
	AddControl ( hDlg, "ltext", 110, 12, 30, 14 , "Proximal:", "string", "");
 	AddControl ( hDlg, "dropdownlist", 145, 10, 60, 70 , sDlgElasticList, "string", "sElasticProximal");

      	Ok = Display("dialog", hDlg, "%center");
  	Close("dialog", hDlg);

	if (Ok == %cancel ) Exit (%ok, "User Cancel!") ;	
	
	if (StringCompare("Selected Point", sZipLength)) EOStoSelectedPoint = %true;
	    
	if (EOStoSelectedPoint )
		Loop = %false ;
	else {
		nZipLength = Value("scalar", sZipLength);
		if ( nZipLength == 0  && StringLength (sZipLength) > 0) {
			Display ("beep", "Invalid given length!") ;
			Loop = %true ;
			}
   		else
			Loop = %false ;
		}
    }   
    
   Display ("cursor", "wait", "Drawing Zipper");
	
   // Check that horizontal zipper line can intersect with the  EOS
   // Get EOS at Proximal End
   EOSFound = %false ;
   sTmp = "type = 'Line' AND layer = 'Construct' AND DB ID = '" + sStyleID + sLeg + "EOSLine'" ;
   hChan = Open ("selection", sTmp );
   if (hChan) {
  	ResetSelection(hChan);
	while (hEOS = GetNextSelection(hChan)) {
	      	GetExtent(hEOS, &xyMin , &xyMax) ;
	      	xyHorizEnd.x = xyMin.x ;  
	      	sTmp = "" ;
	      	GetDBData(hEOS,"Data", sTmp) ;
	     	if (!ScanLine(sTmp, "blank", &nThighFiguredCir, &nWaistActualCir))
   			Exit(%cancel, "Can't Extract actual waist circumferance, Re-Draw Cut Out");	
	    	EOSFound = %true ;	      
		}	
   	}
   Close("selection", hChan); 
   if (!EOSFound) 
   	Exit(%cancel, "Can't find an EOS for End of Zipper at the PROXIMAL EOS!");	
    	
   //Establish allowance for zippers
   //
   nElasticProximal = 0.75 ;
   if (StringCompare(StringMiddle(sElasticProximal,1,1), "N")) nElasticProximal = 0 ;
   if (StringCompare(StringMiddle(sElasticProximal,1,3), "3/8")) nElasticProximal = 0.375 ;
   if (StringCompare(StringMiddle(sElasticProximal,1,1), "1")) nElasticProximal = 1.5 ;
   nElastic = nElasticProximal ;
    
     
// Draw on layer Notes
//
   hLayer = Table("find", "layer", "Notes") ;
   if ( hLayer > %zero )
 	Execute ("menu", "SetLayer",  hLayer) ;
   else
   	Table ("add", "layer", sLayer, "Notes in Green", "current", Table("find", "color", "Green"));

// 
// Draw Zipper
// 
// Establish Zipper Y
//
   nSeam = 0.1875;
   xyHorizEnd.y = xyOtemplate.y + nSeam ;
   xyHorizStart.y = xyHorizEnd.y ;
       	  
   if (EOStoSelectedPoint){
   	xyHorizStart.x = xyHorizEnd.x ;
   	while((xyHorizEnd.x - xyHorizStart.x) <= 0 ){
   		if (!GetUser("xy", "Select Start of Zipper", &xyHorizStart)) Exit(%ok, "Start not selected") ;
		if ((xyHorizEnd.x - xyHorizStart.x) < 0 ) Display("message", "error", "Select towards DISTAL end, Try again");
		}	
   	xyHorizStart.y = xyHorizEnd.y ;
	}
	
   if (nZipLength > 0 )
	xyHorizStart.x = xyHorizEnd.x -  ((nZipLength / 1.2 ) - nElastic) ;
   else	{
   	nDrawnLength = Calc("length",xyHorizStart, xyHorizEnd) ;
	nZipLength = (nDrawnLength + nElastic ) * 1.2 ;
	}
	
// Draw markers only
    	
// Add label and arrows
   AddEntity("marker","closed arrow", xyHorizStart , 0.5 ,0.125, 0 ) ;	
   AddEntity("marker","closed arrow",xyHorizEnd , 0.5 ,0.125, 180 ) ;	

   sZipLength = Format("length", nZipLength);
   sZipOffset = Format("length", nWaistActualCir / 6);

// Add text
// Place relative to xyHorizEnd
   SetData ("TextHorzJust", 1) ;
   SetData ("TextVertJust",32) ;
   SetData ("TextHeight", 0.125) ;
   
   AddEntity ("text",
  		"Place " +  sZipLength + " ZIPPER\n" +
  		sZipOffset + " FROM CENTRE FRONT\n" +
  		"SEAM ON LEFT SIDE"
   		, xyHorizEnd.x - 3, xyHorizEnd.y + 1.25 );
// Reset and exit
   Execute ("menu", "SetLayer",  Table("find", "layer", "1")) ;
   Exit (%ok, "Zipper drawing Complete");

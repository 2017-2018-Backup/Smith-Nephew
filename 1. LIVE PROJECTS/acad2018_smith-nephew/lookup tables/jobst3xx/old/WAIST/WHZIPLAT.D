// File Name:	WHZIPLAT.D
// Purpose:	Draws a Waist Height Lateral body zipper
//		Waist Height Panty Leg
// Version:	1.01
// Date:		09.Mar.1995
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
	hLegBox,
	hCurve,
	hLayer;

XY
	xyInt,		// FN_CirLinInt return value
	xySmallestCir,
	xyConstruct,
	xyStart,
	xyEnd,
	xyCurveStart,
	xyHorizStart,
	xyHorizEnd,
	xyMax, xyMin,
	xyAnkle,
	xyCO_WaistBott, 
	xyOtemplate,
	xyPt1,
	xyTmp;

STRING
	sClass,
	sDlgDistalElasticList, 
	sElasticDistal,
	sDlgProximalElasticList,
	sElasticProximal,
	sMedial,
	sZipperStyle,
	sUnits,
	sZipLength,
	sProfileID,
	sAnkleID,
	sOtemplateID,
	sCO_WaistBottID,
	sTmp, sTmp1,
	sLeg,
	sStyleID,
	sDlgElasticList ,
	sDlgLengthList,
	sLine,
	sTemplate,
	sZipperID,
	sType;

NUMBER
	nMedial,
	nMinZipLength,
	nThighFiguredCir, WaistActualCir,
	nElasticDistal,
	nElasticProximal,
	nElastic,	//sum of Proximal and Distal
	nType,
	ii, nn,
	nZipOff,
	nZipLength ;
BOOL
	EOStoEOS,
	EOStoLegTape,
	EOSFound,
	EOStoSelectedPoint,
	Panty,
	Loop ;
MSGID
	Ok;

// Create procedures 
//
procedure PR_SetZipperID () {
	// Label entity with ID string and make Zipper %true
   	SetDBData(hEnt, "ID", sZipperID );
   	SetDBData(hEnt, "Zipper", "1" );
}

// Find JOBST installed directory
// Set path to macros
//
   STRING	 sPathJOBST;
   sPathJOBST = GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;
   SetData("PathMacro",  sPathJOBST + "\\WAIST" );

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
   sAnkleID        =  sStyleID + sLeg + "Ankle";
   sStyleID = StringMiddle(sStyleID, 5, StringLength(sStyleID)-4) ;

   sCO_WaistBottID = sStyleID + sLeg + "CO_WaistBott";
   sTmp = "type = 'Marker' AND (DB ID ='" + sCO_WaistBottID + "' OR DB ID ='"  + sAnkleID  + "' OR DB ID ='" + sOtemplateID + "')";
   nMarkersFound = 0 ;
   Panty = %true ;
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
   				GetDBData( hEnt, "Pressure", &sTemplate); 	
				GetDBData(hEnt, "units", &sUnits);
				GetDBData(hEnt, "Data", &sTmp) ;
	  			if ( !ScanLine(sTmp, "blank", &nType, &nn)) Exit(%cancel,"Can't get data from Origin Marker!\n");				
				sType = sTmp ;
				}
			if (StringCompare(sTmp, sAnkleID)) {
				Panty = %false ;
				GetGeometry(hEnt, &sTmp, &xyAnkle , &xyTmp, &aAngle) ;
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

// Get Smallest circumference by looping through the vetex of the curve 
// For legs with a foot then use ankle
   if (Panty) { 
   	xySmallestCir.y = 10000 ; // Impossibly large value used as an initial value  
   	nn = 1;
   	while (nn < GetVertexCount(hCurve)) {
        		GetVertex(hCurve, nn, &xyTmp) ;
		if ( xyTmp.y < xySmallestCir.y)
			xySmallestCir.y = xyTmp.y ;
		else
			break ; 
		nn = nn + 1;
		}
	}
    else 
    	xySmallestCir.y = xyAnkle.y ;
    	
// Get EOS at Distal End
// We can assume that this is the start of the Profile
//
   GetVertex(hCurve, 1, &xyCurveStart) ;

// Create Dialog
// Get Zipper style
//
   sUnits = "Inches";

   EOStoSelectedPoint = %false ;
   EOStoEOS = %false;
   EOStoLegTape = %false; 
   sDlgElasticList =  "3/8\" Elastic\n3/4\" Elastic\n1½\" Elastic\nNo Elastic";
   sDlgProximalElasticList =  "3/4\" Elastic\n" + sDlgElasticList;
   sDlgDistalElasticList =  "No Elastic\n" + sDlgElasticList;

   if (Panty)
   	sDlgLengthList =  "Leg Tape\nEOS\nSelected Point";
   else
      	sDlgLengthList =  "Given a length\nSelected Point";

   nMedial = 0 ;
   Loop = %true ;
   nMinZipLength = 5 ;
   while (Loop) {		
            nButX = 65; nButY = 45;
            hDlg = Open ("dialog", sLeg + " Lateral Zipper (Waist Height)", "font Helv 8", 20,20, 210, 65);

	AddControl ( hDlg, "pushbutton",  nButX, nButY, 35, 14, "Cancel", "%cancel", "");
	AddControl ( hDlg, "pushbutton",  nButX + 48 , nButY, 35,14, "OK", "%ok", "");

	AddControl ( hDlg, "ltext", 5, 12, 40, 14 , "EOS to", "string", "");
	AddControl ( hDlg, "combobox", 35, 10, 70, 40 , sDlgLengthList , "string", "sZipLength");
	AddControl ( hDlg, "ltext", 110, 12, 30, 14 , "Proximal:", "string", "");
 	AddControl ( hDlg, "dropdownlist", 145, 10, 60, 70 , sDlgProximalElasticList , "string", "sElasticProximal");
	AddControl ( hDlg, "ltext", 110, 30, 25, 14 , "Distal :", "string", "");
	
 	AddControl ( hDlg, "dropdownlist", 145, 28, 60, 70 , sDlgDistalElasticList , "string", "sElasticDistal");

   	Ok = Display("dialog", hDlg, "%center");
  	Close("dialog", hDlg);

	if (Ok == %cancel ) Exit (%ok, "User Cancel!") ;	
	
	if (StringCompare("Leg Tape", sZipLength)) EOStoLegTape = %true;
	if (StringCompare("EOS", sZipLength)) EOStoEOS = %true;
	if (StringCompare("Selected Point", sZipLength)) EOStoSelectedPoint = %true;
	    
	if (EOStoLegTape || EOStoEOS || EOStoSelectedPoint )
		Loop = %false ;
	else {
		nZipLength = Value("scalar", sZipLength);
		if ( nZipLength == 0 && StringLength (sZipLength) > 0) {
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
   			Exit(%cancel, "Can't Extract Figured Thigh Circumference or Actual Waist, Re-Draw Cut Out");	
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

     if (EOStoSelectedPoint || EOStoLegTape || nZipLength > 0 ) //closed zipper at distal end
     	nElasticDistal = 0 ;
     else	{
     	nElasticDistal = 0.75 ;
     	if (StringCompare(StringMiddle(sElasticDistal,1,1), "N")) nElasticDistal = 0 ;
     	if (StringCompare(StringMiddle(sElasticDistal,1,3), "3/8")) nElasticDistal = 0.375 ;
     	if (StringCompare(StringMiddle(sElasticDistal,1,1), "1")) nElasticDistal = 1.5 ;
     	}

    nElastic = nElasticDistal + nElasticProximal ;
    
// Get template from leg box
// To establish Tape spacing 

   nTemplate = Value ("scalar", StringMiddle(sTemplate,1,2));
   if (nTemplate >= 30) nTapeSpacing = 1.25 ;
   if (nTemplate == 13) nTapeSpacing = 1.31 ;  
   if (nTemplate == 9) nTapeSpacing= 1.37  ;  
                       
     
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
// NB nThighFiguredCir is given using the 1/2 scale.
   nZipOff = 1.125;
   if (Panty)
  	xyHorizEnd.y = xyOtemplate.y + ((nThighFiguredCir * 2 )/ 6) ;
   else
  	xyHorizEnd.y = xyOtemplate.y + (nWaistActualCir / 6) ;
   
   if (xyHorizEnd.y > xySmallestCir.y - nZipOff ) 
   	xyHorizEnd.y =  xySmallestCir.y - nZipOff  ;
   xyHorizStart.y = xyHorizEnd.y ;
       	
// For Medial Zipper step back 3 tapes from groin height
// Establish X of start and end

   if (EOStoLegTape)
   	xyHorizStart.x = xyCurveStart.x + nTapeSpacing ;

   if (EOStoEOS)
   	xyHorizStart.x = xyCurveStart.x;
  
   if (EOStoSelectedPoint){
   	xyHorizStart.x = xyHorizEnd.x ;
   	while((xyHorizEnd.x - xyHorizStart.x) <= 0){
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
	
// Draw Zip
   SetSymbolLibrary( sPathJOBST + "\\JOBST.SLB") ;   	     
   Symbol( "find", "TextAsSymbol") ;
   hEnt = AddEntity ("symbol", "TextAsSymbol", xyHorizStart.x + (xyHorizEnd.x- xyHorizStart.x)/2, xyHorizStart.y + .25, 1, 1, 0) ;  

// Create ID string
   sZipperID =  sStyleID + MakeString("scalar", UID("get", hEnt )) ;
   SetDBData(hEnt, "ID", sZipperID );
   SetDBData(hEnt, "Zipper", "1" );
   SetDBData(hEnt, "Data",  Format("length",nZipLength) +  " LATERAL ZIPPER" );

// Draw zip line
   hEnt = AddEntity("line",xyHorizStart ,xyHorizEnd) ;
   PR_SetZipperID() ;
 	
// Add label and arrows
   hEnt = AddEntity("marker","closed arrow", xyHorizStart , 0.5 ,0.125, 0 ) ;	
   PR_SetZipperID() ;
   if (     xyHorizEnd.y > xyCO_WaistBott.y 
       && xyCO_WaistBott.y != 0 
       && xyCO_WaistBott.x != 0)
   	hEnt = AddEntity("marker","closed arrow",xyHorizEnd.x, xyCO_WaistBott.y - 0.025, 0.5 ,0.125, 225 ) ;	
    else 
   	hEnt = AddEntity("marker","closed arrow",xyHorizEnd , 0.5 ,0.125, 180 ) ;	
    PR_SetZipperID() ;

 // Reset and exit
   Execute ("menu", "SetLayer",  Table("find", "layer", "1")) ;
   Exit (%ok, "Zipper drawing Complete");

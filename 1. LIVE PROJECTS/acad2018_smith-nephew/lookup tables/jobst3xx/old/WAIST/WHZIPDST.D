// File Name:	WHZIPDST.D
// Purpose:	Draws a Zip from DISTAL EOS up
//			FootLess styles only
//			Waist Heights
// Version:	1.01
// Date:		15.Dec.1994
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action

HANDLE
	hDlg,
	hEnt,
	hCurve,
	hEOS,
	hChan,
	hLeg,
	hLayer;

XY
	xyInt,		// FN_CirLinInt return value
	xySmallestCir,
	xyAnkle,
	xyConstruct,
	xyStart,
	xyEnd,
	xyHorizStart,
	xyHorizEnd,
	xyEOSCen,
	xyMax, xyMin,
	xyOtemplate,
	xyPt1,
	xyTmp;

STRING
	sClass,
	sElastic,
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
	sTmp, sTmp1,
	sLeg,
	sStyleID,
	sDlgElasticList ,
	sDlgLengthList,
	sType;

NUMBER
	nMedial,
	nElastic,
	nType,
	ii, nn,
	nY,nX,
	nEOSRad,
	nZipLength,
	nHeelOff,
	nHeelRad;

BOOL
	EOStoEOS,
	EOSFound,
	EOStoSelectedPoint,
	Loop,
	Panty;
	
ANGLE
	aAngle;

MSGID
	Ok;
// Create procedures 
//
   STRING	  sZipperID;
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
   GetDBData (hCurve , "ID", &sProfileID);
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
   sTmp = "type = 'Marker' AND (DB ID ='"  + sAnkleID  + "' OR DB ID ='" + sOtemplateID + "')";
   Panty = %true ;
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
				sType = sTmp ;
				}
			if (StringCompare(sTmp, sAnkleID)) {
				Panty = %false ;
				GetGeometry(hEnt, &sTmp, &xyAnkle , &xyTmp, &aAngle) ;
				}

			}
		}
   	}
   Close("selection", hChan); 

// Check if that the markers have been found, otherwise exit
//
    if (StringCompare("CHAP", sType))Exit (%cancel, "Can't use this zipper for a CHAP style!\n") ;
    if (Panty && nMarkersFound < 1) Exit (%cancel, "Missing markers for selected style, data not found!\n") ;
    if (Panty && nMarkersFound > 1) Exit (%cancel, "Two or more drawings of the same style exist!\nDelete the extra drawing/s and try again.\n") ;
    if (!Panty) Exit (%cancel, "Can only be used on Footless styles!") ;
     
// Get Smallest circumference 
// Get Smallest circumference by looping through the vetex of the curve 
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
   

// Create Dialog
//
   sUnits = "Inches";
   EOStoSelectedPoint = %false ;
   
   Loop = %true ;
   nMinZipLength = 5 ;
   
   sDlgElasticList =  "3/8\" Elastic\n3/4\" Elastic\n1½\" Elastic\nNo Elastic";
   	
   sDlgProximalElasticList =  "No Elastic\n" + sDlgElasticList;
   sDlgDistalElasticList =  "No Elastic\n" + sDlgElasticList;

   sDlgLengthList =  "Give a length\nSelected Point";   
         
   nMedial = 0 ;
   while (Loop) {		
            nButX = 70; nButY = 65;
            hDlg = Open ("dialog", sLeg + " Distal EOS Zipper (Waist Height)", "font Helv 8", 20,20, 235, 95);

	AddControl ( hDlg, "pushbutton",  nButX, nButY, 35, 14, "Cancel", "%cancel", "");
	AddControl ( hDlg, "pushbutton",  nButX + 48 , nButY, 35,14, "OK", "%ok", "");

	AddControl ( hDlg, "ltext", 5, 12, 55, 14 , "Distal EOS to", "string", "");
	AddControl ( hDlg, "combobox", 60, 10, 70, 40 , sDlgLengthList , "string", "sZipLength");

	AddControl ( hDlg, "ltext", 135, 12, 25, 14 , "Distal :", "string", "");
	AddControl ( hDlg, "dropdownlist", 170, 10, 60, 70 , sDlgDistalElasticList , "string", "sElasticDistal");

	AddControl ( hDlg, "ltext", 135, 30, 30, 14 , "Proximal:", "string", "");
 	AddControl ( hDlg, "dropdownlist", 170, 28, 60, 70 , sDlgProximalElasticList , "string", "sElasticProximal");
 	

  	AddControl ( hDlg, "checkbox", nButX + 12, nButY - 18, 65, 15 , "Medial Zipper", "number", "nMedial");

   	Ok = Display("dialog", hDlg, "%center");
  	Close("dialog", hDlg);

	if (Ok == %cancel ) Exit (%ok, "User Cancel!") ;	

	if (StringCompare("Selected Point", sZipLength)) EOStoSelectedPoint = %true;
	    
	if (EOStoEOS || EOStoSelectedPoint )
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

// Start X and Y
//
    nZipOff = 1.125 ;
    GetVertex(hCurve,1, &xyHorizStart) ;
    xyHorizStart.y = xySmallestCir.y - nZipOff  ;	
    xyHorizEnd.y = xyHorizStart.y ;

// Establish allowance for zippers
//
     nElasticProximal = 0.75 ;
     if (StringCompare(StringMiddle(sElasticProximal,1,1), "N")) nElasticProximal = 0 ;
     if (StringCompare(StringMiddle(sElasticProximal,1,3), "3/8")) nElasticProximal = 0.375 ;
     if (StringCompare(StringMiddle(sElasticProximal,1,1), "1")) nElasticProximal = 1.5 ;

     if (EOStoSelectedPoint || nZipLength > 0 ) //closed zipper at distal end
     	nElasticDistal = 0 ;
     else	{
     	nElasticDistal = 0.75 ;
     	if (StringCompare(StringMiddle(sElasticDistal,1,1), "N")) nElasticDistal = 0 ;
     	if (StringCompare(StringMiddle(sElasticDistal,1,3), "3/8")) nElasticDistal = 0.375 ;
     	if (StringCompare(StringMiddle(sElasticDistal,1,1), "1")) nElasticDistal = 1.5 ;
     	}

    nElastic = nElasticDistal + nElasticProximal ;

// Draw on layer notes
//
   hLayer = Table("find", "layer", "Notes") ;
   if ( hLayer > %zero )
 	Execute ("menu", "SetLayer",  hLayer) ;
   else
   	Table ("add", "layer", sLayer, "Notes in Green", "current", Table("find", "color", "Green"));

   if (EOStoSelectedPoint){
   	xyHorizEnd.x = xyHorizStart.x  ;
   	while((xyHorizEnd.x - xyHorizStart.x) <= 0 ){
   		if (!GetUser("xy", "Select End of Zipper", &xyHorizEnd)) Exit(%ok, "End not selected") ;
			if ((xyHorizEnd.x - xyHorizStart.x) < 0 ) Display("message", "error", "Select towards PROXIMAL end, Try again");
		}	
   	xyHorizEnd.y = xyHorizStart.y ;
	}
	
   if (nZipLength > 0 )
	xyHorizEnd.x = xyHorizStart.x + ((nZipLength / 1.2 ) - nElastic) ;
   else	{
   	nDrawnLength = Calc("length",xyHorizStart, xyHorizEnd) ;
	nZipLength = (nDrawnLength + nElastic ) * 1.2 ;
	}
	   
// Add zipper    	
   if (nMedial == 1 )
   	sTmp = Format("length",nZipLength) +  " MEDIAL ZIPPER" ;
   else
      	sTmp =  Format("length",nZipLength) +  " LATERAL ZIPPER" ;
      	
// Draw Zip text 
   SetSymbolLibrary( sPathJOBST + "\\JOBST.SLB") ;   	     
   Symbol( "find", "TextAsSymbol") ;
   hEnt = AddEntity ("symbol", "TextAsSymbol", xyHorizStart.x + (xyHorizEnd.x- xyHorizStart.x)/2, xyHorizStart.y + .25, 1, 1, 0) ;  

// Create ID string
   sZipperID =  sStyleID + MakeString("scalar", UID("get", hEnt )) ;
   SetDBData(hEnt, "ID", sZipperID );
   SetDBData(hEnt, "Zipper", "1" );
   SetDBData(hEnt, "Data", sTmp );

// Draw zip line
   hEnt = AddEntity("line",xyHorizStart ,xyHorizEnd) ;
   PR_SetZipperID() ;
  
// Add label and arrows
   hEnt = AddEntity("marker","closed arrow", xyHorizStart , 0.5 ,0.125, 0 ) ;	
   PR_SetZipperID() ;

   hEnt = AddEntity("marker","closed arrow",xyHorizEnd , 0.5 ,0.125, 180 ) ;	
   PR_SetZipperID() ;

// Reset and exit
   Execute ("menu", "SetLayer",  Table("find", "layer", "1")) ;
   Exit (%ok, "Zipper drawing Complete");

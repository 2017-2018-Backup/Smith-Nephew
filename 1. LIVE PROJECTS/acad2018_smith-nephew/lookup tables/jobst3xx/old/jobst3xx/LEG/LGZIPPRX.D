// File Name:	LGZIPPRX.D
// Purpose:	Draws a Zip from PROXIMAL EOS down
//		Lower Leg
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
	xyConstruct,
	xyStart,
	xyEnd,
	xyCurveStart,
	xyHorizStart,
	xyHorizEnd,
	xyEOSCen,
	xyMax, xyMin,
	xyAnkle,
	xyHeel,
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
	sHeelID,
	sAnkleID,
	sOtemplateID,
	sTmp, sTmp1,
	sLeg,
	sStyleID,
	sDlgElasticList ,
	sDlgLengthList,
	sLine,
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
	FootLess,
	SmallHeel;

MSGID
	Ok;
// Reset 
    UserSelection ("clear");
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
   sHeelID          = sStyleID + sLeg+ "Heel";
   sOtemplateID = sStyleID + sLeg + "Origin";
   sAnkleID        =  sStyleID + sLeg + "Ankle";
   sTmp = "type = 'Marker' AND (DB ID ='" + sHeelID + "' OR DB ID ='"  + sAnkleID  + "' OR DB ID ='" + sOtemplateID + "')";
   nMarkersFound = 0 ;
   FootLess = %true ;
   hChan = Open ("selection", sTmp );
   if (hChan) {
	ResetSelection(hChan);
	while (hEnt = GetNextSelection(hChan)) {
		GetDBData(hEnt,"ID", &sTmp) ;
		GetGeometry(hEnt, &sTmp1, &xyPt1, &xyTmp, &aAngle) ;
	                if ( xyPt1.x != 0 || xyPt1.y != 0) {
			if (StringCompare(sTmp, sHeelID )){
				nMarkersFound = nMarkersFound + 1 ;
				GetGeometry(hEnt, &sTmp, &xyHeel, &xyTmp, &aAngle) ;
				GetDBData(hEnt, "Data", &sTmp);
				SmallHeel = Value("scalar", sTmp);
				}
			if (StringCompare(sTmp, sOtemplateID)) {
				nMarkersFound = nMarkersFound + 1 ;
				GetGeometry(hEnt, &sTmp, &xyOtemplate , &xyTmp, &aAngle) ;
				GetDBData(hEnt, "units", &sUnits);
				GetDBData(hEnt, "Data", &sTmp) ;
	  			if ( !ScanLine(sTmp, "blank", &nType, &nn)) Exit(%cancel,"Can't get data from Origin Marker!\n");				
				}
			if (StringCompare(sTmp, sAnkleID)) {
				nMarkersFound = nMarkersFound + 1 ;
				FootLess = %false;
				GetGeometry(hEnt, &sTmp, &xyAnkle , &xyTmp, &aAngle) ;
				GetDBData(hEnt, "Data", &sTmp);
				if ( !ScanLine(sTmp, "blank", &nX, &nY)) Exit(%cancel,"Can't get data from Ankle Marker!\n");
				}
			}
		}
   	}
   Close("selection", hChan); 

// Check if that the markers have been found, otherwise exit
//
    if (!FootLess && nMarkersFound < 3) Exit (%cancel, "Missing markers for selected foot, data not found!\n") ;
    if (!FootLess && nMarkersFound > 3) Exit (%cancel, "Two or more drawings of the same style exist!\nDelete the extra drawing/s and try again.\n") ;

// Get Smallest circumference (by requesting) a point
// Do some checking that the given point is sensible
   Loop = %true ;	
   // Use First vertex for footless style
   if (FootLess)
   	GetVertex (hCurve , 1, &xySmallestCir) ;
   else
         	while (Loop) {
   		GetUser("xy", "Select smallest leg circumference", &xySmallestCir);
		if ( xySmallestCir.y > xyAnkle.y || xySmallestCir.y < xyOtemplate.y){
			Display ("beep", "Selected point above Ankle | Curve Start or below edge of template. Try again");
			Loop = %true;
			 }
		else 
			Loop = %false;
		}

// Create Dialog
// Get Zipper style
//
   sUnits = "Inches";
   EOStoSelectedPoint = %false ;
   EOStoEOS = %false;

   Loop = %true ;
   nMinZipLength = 5 ;
   
   sDlgElasticList =  "3/8\" Elastic\n3/4\" Elastic\n1½\" Elastic\nNo Elastic";
   	
   sDlgProximalElasticList =  "3/4\" Elastic\n" + sDlgElasticList;
   sDlgDistalElasticList =  "No Elastic\n" + sDlgElasticList;

   if (FootLess) {
         if (nType == 3 || nType == 4 ) 
         	sDlgDistalElasticList =   "3/4\" Elastic\n" + sDlgElasticList; // Knee Band && Thigh Band A/K
         if (nType == 5)  
         	sDlgDistalElasticList =   "No Elastic\n" + sDlgElasticList;  // Thigh Band B/K
         sDlgLengthList =  "Give a length\nSelected Point\nDistal EOS";   
         }
   else
         sDlgLengthList =  "Give a length\nSelected Point";
         
   nMedial = 0 ;
   while (Loop) {		
            nButX = 70; nButY = 65;
            hDlg = Open ("dialog", sLeg + " Proximal EOS Zipper (Lower Extremity)", "font Helv 8", 20,20, 235, 95);

	AddControl ( hDlg, "pushbutton",  nButX, nButY, 35, 14, "Cancel", "%cancel", "");
	AddControl ( hDlg, "pushbutton",  nButX + 48 , nButY, 35,14, "OK", "%ok", "");

	AddControl ( hDlg, "ltext", 5, 12, 55, 14 , "Proximal EOS to", "string", "");
	AddControl ( hDlg, "combobox", 60, 10, 70, 40 , sDlgLengthList , "string", "sZipLength");
	AddControl ( hDlg, "ltext", 135, 12, 30, 14 , "Proximal:", "string", "");
 	AddControl ( hDlg, "dropdownlist", 170, 10, 60, 70 , sDlgProximalElasticList , "string", "sElasticProximal");
	AddControl ( hDlg, "ltext", 135, 30, 25, 14 , "Distal :", "string", "");
	
 	AddControl ( hDlg, "dropdownlist", 170, 28, 60, 70 , sDlgDistalElasticList , "string", "sElasticDistal");

  	AddControl ( hDlg, "checkbox", nButX + 12, nButY - 18, 65, 15 , "Medial Zipper", "number", "nMedial");

   	Ok = Display("dialog", hDlg, "%center");
  	Close("dialog", hDlg);

	if (Ok == %cancel ) Exit (%ok, "User Cancel!") ;	
	
	if (StringCompare("Distal EOS", sZipLength)) EOStoEOS = %true;
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

   Display ("cursor", "wait", "Drawing Zipper");
	
   nZipOff = 1.125 ;
   xyHorizStart.x =  xySmallestCir.x ; 
   xyHorizStart.y = xySmallestCir.y - nZipOff  ;	
	
   // Check that horizontal zipper line can intersect with the  EOS
          EOSFound = %false ;
          sTmp = "(type = 'Arc' OR type = 'Line') AND DB ID = '" + sStyleID + sLeg + "ClosingLine'" ;
          hChan = Open ("selection", sTmp );
          if (hChan) {
	ResetSelection(hChan);
	while (hEOS = GetNextSelection(hChan)) {
	      GetEntityClass(hEOS, &sClass);
	      GetExtent(hEOS, &xyMin , &xyMax) ;
	      if (xyHorizStart.y >= xyMin.y && xyHorizStart.y <= xyMax.y) {
		if (StringCompare(sClass, "arc")) {
			GetGeometry( hEOS, &xyEOSCen, &nEOSRad, &aStart, &aDelta);
			xyHorizEnd.x =  xyEOSCen.x + sqrt ( nEOSRad^2 - ( xyHorizStart.y - xyEOSCen.y)^2);
			}
		else	{
			if (xyMin.x == xyMax.x)  
				xyHorizEnd.x = xyMin.x ;  
			else {	
				GetGeometry(hEOS, &xyStart, &xyEnd);
				nTanTheta = (xyEnd.y - xyStart.y) / ( xyEnd.x - xyStart.x);
				xyHorizEnd.x = xyMin.x  +  (xyHorizStart.y - xyMin.y) / nTanTheta ;
				}

			}
		EOSFound = %true ;
		xyHorizEnd.y = xyHorizStart.y ;
		}
	     }	
	}
	Close("selection", hChan); 
   	if (!EOSFound) {
   		Display("message","error", "Can't find an EOS for zipper line to intersect!\nSelect a point manually!");	
   		Select = %true ;
   		}

   //Establish allowance for zippers
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

// Establish X
   if (EOStoEOS) {
   	 // Only applicable to footless
   	GetVertex(hCurve, 1 , &xyTmp) ; 
  	xyHorizStart.x = xyTmp.x;
  	}
   
   if (EOStoSelectedPoint){
   	xyHorizStart.x = xyHorizEnd.x ;
   	while((xyHorizEnd.x - xyHorizStart.x) <= 0 ){
   		if (!GetUser("xy", "Select Start of Zipper", &xyHorizStart)) Exit(%ok, "Start not selected") ;
		if ((xyHorizEnd.x - xyHorizStart.x) < 0 ) Display("message", "error", "Select towards DISTAL end, Try again");
		}	
   	xyHorizStart.y = xyHorizEnd.y ;
	}

// draw Zip
//
   if (nZipLength > 0 )
	xyHorizStart.x = xyHorizEnd.x -  ((nZipLength / 1.2 ) - nElastic) ;
	if (xyHorizStart.x < xySmallestCir.x && !FootLess) {
   		nArcLength = (xySmallestCir.x - xyHorizStart.x) ;
   		if (nArcLength < .125)
   			DrawStraight = %true;
   		else	{
   			DrawStraight = %false;
   			// curve end of line towards heel
   			if (SmallHeel) 
				nHeelRad = 1.1760709 ; 
			else
				nHeelRad = 1.5053069 ;
			nRadius = nHeelRad + 1.125 ;
			xyConstruct.x =  xyAnkle.x + nX;  
   			xyConstruct.y =  xyAnkle.y + nY;  // nY from ankle 
			xyHorizStart.x = xySmallestCir.x ;
			aAngle = Calc("angle",xyConstruct, xyHorizStart) ;
			aPrevAngle = aAngle - ((nArcLength*180) / (%pi*nRadius)) ;
			aAngleInc = (aAngle - aPrevAngle) / 4 ;
			StartPoly("fitted") ;
			 ii = 0 ;
  			 while ( ii  <= 3) {	//*
			 	xyTmp = CalcXY ("relpolar", xyConstruct , nRadius, aPrevAngle + aAngleInc * ii); //*
			 	if (xyTmp.x < xyHorizStart.x) AddVertex ( xyTmp); //*
			 	ii = ii + 1 ;
				}
			 AddVertex ( xyHorizStart ) ;
			 AddVertex ( xyHorizStart.x + 0.25, xyHorizStart.y) ;
			 AddVertex ( xyHorizStart.x + 0.5, xyHorizStart.y) ;
			 AddVertex ( xyHorizEnd ) ;
			EndPoly() ;
			AddEntity("marker","closed arrow",  CalcXY ("relpolar", xyConstruct , nRadius, aPrevAngle) , 0.5 ,0.125, aPrevAngle + 92 ) ;	
   			}   		
		}
	
   else	{
   	DrawStraight = %true ;
   	nDrawnLength = Calc("length",xyHorizStart, xyHorizEnd) ;
	nZipLength = (nDrawnLength + nElastic ) * 1.2 ;
	}

// Draw Straight Zip
//
   if (DrawStraight){
	// Draw straight line as normal
	AddEntity("line",xyHorizStart ,xyHorizEnd) ;
	AddEntity("marker","closed arrow", xyHorizStart , 0.5 ,0.125, 0 ) ;	
	}

   
    	
// Add label and arrows
   AddEntity("marker","closed arrow",xyHorizEnd , 0.5 ,0.125, 180 ) ;	
   if (nMedial == 1 )
   	sTmp = Format("length",nZipLength) +  " MEDIAL ZIPPER" ;
   else
      	sTmp =  Format("length",nZipLength) +  " LATERAL ZIPPER" ;
   SetData ("TextVertJust", 32) ;
   SetData ("TextHorzJust", 2) ;   	
   SetData ("TextHeight", 0.125) ;

   AddEntity ("text",sTmp, xyHorizStart.x + (xyHorizEnd.x- xyHorizStart.x)/2, xyHorizStart.y + .25);

// Reset and exit
   Execute ("menu", "SetLayer",  Table("find", "layer", "1")) ;
   Exit (%ok, "Zipper drawing Complete");

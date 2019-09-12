// File Name:	WHZIPCHP.D
// Purpose:	Draws an Ankle Zip etc for
//		CHAP style Leg
// Version:	1.01
// Date:		4.Oct.1996
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action

HANDLE
	hDlg,
	hEnt,
	hChan,
	hTextEnt,
	hLayerNotes,
	hEntClosure,
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
	xyFold,
	xyEOS,
	xyZipperEndInChap,
	xyZipperStartInChap,
	xyOtemplate,
	xyPt1,
	xyTmp;

STRING
	sClass,
	sZipperID,
	sElastic,
	sMedial,
	sZipperStyle,
	sUnits,
	sZipLength,
	sProfileID,
	sHeelID,
	sAnkleID,
	sFoldID,
	sEOSID,
	sOtemplateID,
	sCO_WaistBottID,
	sTmp, sTmp1,
	sLeg,
	sStyleID,
	sDlgElasticList ,
	sDlgLengthList,
	sLine,
	sTemplate,
	sType;

NUMBER
	nMedial,
	nZipperInChap,
	nElastic,
	nType,
	ii, nn,
	nY,nX,
	nEOSRad,
	nZipLength,
	nHeelOff,
	nHeelRad;

BOOL
	EOS,
	EOSFound,
	DrawCurve,
	Select,
	Loop,
	SmallHeel;

MSGID
	Ok;

ANGLE
	aStart,
	aDelta,
	aPrevAngle,
	aAngle;
	
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

// Sub Routines etc
//
   @WHZIPDEF.D;

// Get profile, identify Leg and FileNo
//
   if(!hEnt = GetEntity ("Select a Leg Profile", "curve")) Exit (%ok, "User cancelled") ;

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
   sStyleID         = sStyleID + sLeg ;
   sHeelID          = sStyleID + "Heel";
   sOtemplateID = sStyleID + "Origin";
   sAnkleID         = sStyleID +  "Ankle";
   sFoldID	          = sStyleID  + "Fold";
   sEOSID	          = sStyleID  + "EOS";
   sTmp = "type = 'Marker' AND (DB ID ='" + sHeelID + "' OR DB ID ='"  + sAnkleID  + "'OR DB ID ='"  + sFoldID  + "'OR DB ID ='"  + sEOSID  + "' OR DB ID ='"  + sCO_WaistBottID + "' OR DB ID ='" + sOtemplateID + "')";
   nMarkersFound = 0 ;
   nMarkersRequired = 5 ;
   hChan = Open ("selection", sTmp );
   if (hChan) {
	ResetSelection(hChan);
	while (hEnt = GetNextSelection(hChan)) {
		GetDBData(hEnt,"ID", &sTmp) ;
		GetGeometry(hEnt, &sTmp1, &xyPt1, &xyTmp, &aAngle) ;
	            if ( xyPt1.x != 0 || xyPt1.y != 0) {
			if (StringCompare(sTmp, sHeelID )){
				nMarkersFound = nMarkersFound + 1 ;
				xyHeel = xyPt1 ;
				GetDBData(hEnt, "Data", &sTmp);
				SmallHeel = Value("scalar", sTmp);
				}
			if (StringCompare(sTmp, sOtemplateID)) {
				nMarkersFound = nMarkersFound + 1 ;
				xyOtemplate = xyPt1 ;
				GetDBData(hEnt, "units", &sUnits);
				GetDBData(hEnt, "Pressure", &sTemplate);
   				GetDBData(hEnt, "Data", &sType); 	
				}
			if (StringCompare(sTmp, sAnkleID)) {
				nMarkersFound = nMarkersFound + 1 ;
				xyAnkle = xyPt1 ;
				GetDBData(hEnt, "Data", &sTmp);
				if ( !ScanLine(sTmp, "blank", &nX, &nY)) Exit(%cancel,"Can't get data from Ankle Marker!\n");
				}
			if (StringCompare(sTmp, sFoldID )){
				nMarkersFound = nMarkersFound + 1 ;
				xyFold = xyPt1 ;
				}
			if (StringCompare(sTmp, sEOSID )){
				nMarkersFound = nMarkersFound + 1 ;
				xyEOS = xyPt1 ;
				}
			}
		}
   	}
   Close("selection", hChan); 

// Check if that the markers have been found, otherwise exit
//
    if (!StringCompare("CHAP", sType))Exit (%cancel, "Can't use this zipper for a Waist Ht style!\n") ;
    if (nMarkersFound < nMarkersRequired ) Exit (%cancel, "Missing markers for selected leg profile, data not found!\nOr the leg does not have a foot!") ;
    if (nMarkersFound > nMarkersRequired ) Exit (%cancel, "Two or more drawings of the same style exist!\nDelete the extra drawing/s and try again.\n") ;

// Get Smallest circumference (by requesting) a point
// Do some checking that the given point is sensible
   Loop = %true ;	
   while (Loop) {
   	GetUser("xy", "Select smallest leg circumference", &xySmallestCir);
	if ( xySmallestCir.y > xyAnkle.y || xySmallestCir.y < xyOtemplate.y){
		Display ("beep", "Selected point above Ankle or below edge of template. Try again");
		Loop = %true;
		 }
	else 
		Loop = %false;
	}

// Create Dialog
// Get Zipper style
//
   sUnits = "Inches";
   EOS = %false;
   Loop = %true ;
   Select = %false ;
   sDlgElasticList =  "3/8\" Elastic\n3/4\" Elastic\n1½\" Elastic\nNo Elastic";
   sDlgElasticList =  "3/4\" Elastic\n" + sDlgElasticList;
   sDlgLengthList =  "EOS\nGive a length\nSelected Point\n";
   nMedial = 0 ;
   while (Loop) {		
        nButX = 60; nButY = 45;
        hDlg = Open ("dialog", sLeg + "Ankle Zipper (CHAP)", "font Helv 8", 20,20, 215, 70);

	AddControl ( hDlg, "pushbutton",  nButX, nButY, 35, 14, "Cancel", "%cancel", "");
	AddControl ( hDlg, "pushbutton",  nButX + 48 , nButY, 35,14, "OK", "%ok", "");

	AddControl ( hDlg, "ltext", 5, 12, 30, 14 , "Ankle to", "string", "");
	AddControl ( hDlg, "combobox", 40, 10, 65, 45 , sDlgLengthList , "string", "sZipLength");
	AddControl ( hDlg, "ltext", 113, 12, 33, 14 , "Proximal:", "string", "");
 	AddControl ( hDlg, "dropdownlist", 145, 10, 60, 60 , sDlgElasticList, "string", "sElastic");
  	AddControl ( hDlg, "ltext", 5 , nButY - 15, 60, 15 ,"Template: " + sTemplate, "String", "");
  	AddControl ( hDlg, "checkbox", 120, nButY - 18, 60, 15 , "Medial Zipper", "number", "nMedial");
  	
   	Ok = Display("dialog", hDlg, "%center");
  	Close("dialog", hDlg);

	if (Ok == %cancel ) Exit (%ok, "User Cancel!") ;	
	
	if (StringCompare("EOS", sZipLength) || StringCompare("Selected Point", sZipLength)){ 
		Loop = %false ;
		if (StringCompare("EOS", sZipLength)) EOS = %true ;
			else Select = %true ;
		}
	else {
		nZipLength = Value("scalar", sZipLength);
		if ( nZipLength == 0 && StringLength (sZipLength) > 0) 
			Loop = %true ;
   		else {
			Loop = %false ;
			EOS = %false;
			Select = %false;
		        }
		}
    }   

   Display ("cursor", "wait", "Drawing Zipper");

// Draw Chosen Zipper style
// 
   if (SmallHeel) {
	nHeelRad = 1.1760709 ; 
	nHeelOff = 0.25 ;
	}
   else {
	nHeelRad = 1.5053069 ;
	nHeelOff = 0.5;
	}
   nZipOff = 1.125 ;
	
   xyConstruct.x =  xyAnkle.x + nX ; // nX and nY stored relative position of the 
   xyConstruct.y =  xyAnkle.y + nY;  // heel construction circle center point 

   Ok = FN_CirLinInt ( xyHeel.x + nHeelOff , xyConstruct.y, xyHeel.x + nHeelOff, 0, xyConstruct, nHeelRad + nZipOff);
   if (Ok) {
   	// Draw polyline curve from intersection
      	xyCurveStart = xyInt;
   	xyHorizStart.x =  xySmallestCir.x ; 
  	xyHorizStart.y = xySmallestCir.y - nZipOff  ;	
  	DrawCurve = %true;
  	}
   else	{
   	// Otherwise degenerate to a straight line
   	xyHorizStart.x = xyHeel.x + nHeelOff;
   	xyHorizStart.y = xyHeel.y - nZipOff  ;	
   	DrawCurve = %false;
   	}
	
    nZipperInChap  = 0;	
   //if  EOS given
    if (EOS) {
          if (nMedial == 1 ) {
	 	 xyHorizEnd.x = xyFold.x ;  
   		 nElastic = 0 ;
   		 nZipperInChap = 0 ;
		}
          else{
          	xyZipperStartInChap.x = xyFold.x ;
          	xyZipperStartInChap.y = xyFold.y - 1.125 ;
          	xyZipperEndInChap = xyEOS ;
          	xyZipperEndInChap.y = xyFold.y - 1.125;
          	nZipperInChap = Calc("length", xyZipperStartInChap, xyZipperEndInChap) ;        	
          	xyHorizEnd.x = xyFold.x ;    
   		// Elastic allowance at EOS
		nElastic = 0.75 ;
		if (StringCompare(StringMiddle(sElastic,1,1), "N")) nElastic = 0 ;
		if (StringCompare(StringMiddle(sElastic,1,3), "3/8")) nElastic = 0.375 ;
		if (StringCompare(StringMiddle(sElastic,1,1), "1")) nElastic = 1.5 ;
		}  
	  xyHorizEnd.y = xyHorizStart.y ;      
          } 
   else {
   	// No elastic as this implies a closed zipper
   	nElastic = 0 ;
   	//create a dummy xyHorizEnd for use on insertion
   	if (xyHorizStart.x + (nZipLength /1.2) > xyFold.x) 
   		xyHorizEnd.x = xyFold.x ;  
   	else
   		xyHorizEnd.x = xyHorizStart.x + (nZipLength / 1.2)  ;
   	xyHorizEnd.y = xyHorizStart.y ;  
   	}
   	
// For Medial Zipper step back 3 tape
   if (nMedial == 1) {
   	nFoldOffset = 0 ;
   	nTemplate = Value ("scalar", StringMiddle(sTemplate,1,2));
   	if (nTemplate >= 30) nFoldOffset = 1.25 * 3 ;
   	if (nTemplate == 13){
   		nFoldOffset = 1.31 * 3 ;  
   		Display("message", "error", "WARNING!\nTemplate drawn using 13 Down Stretch!");
   		}
   	if (nTemplate == 9)nFoldOffset = 1.37 * 3 ;  
   	xyHorizEnd.x = xyHorizEnd.x - nFoldOffset ;
   	}
                       
// If select given then prompt user for end of zip
//
   if (Select) {
   	if (!GetUser("xy", "Select End of Zipper", &xyHorizEnd)) Exit(%ok, "End not selected") ;
   	xyHorizEnd.y = xyHorizStart.y ;
   	if (xyHorizEnd.x > xyFold.x ) {
          	xyZipperStartInChap.x = xyFold.x ;
          	xyZipperStartInChap.y = xyFold.y - 1.125;
          	xyZipperEndInChap.x = xyHorizEnd.x ;
          	xyZipperEndInChap.y = xyFold.y - 1.125;
          	if (xyZipperEndInChap.x < xyFold.x) 
          		nZipperInChap = 0 ;
          	else
          		nZipperInChap = Calc("length", xyZipperStartInChap, xyZipperEndInChap) ;        	
   		xyHorizEnd.x = xyFold.x ;
   	   	nElastic = 0 ;
		EOS = %true;
   		}
        }
     
// Revise radius to include Zip offset
//
   nHeelRad = nHeelRad + nZipOff ;
   
   hLayerNotes = Table("find", "layer", "Notes") ;
   if ( hLayerNotes <> %badtable )
 	Execute ("menu", "SetLayer",  hLayerNotes );
   else
   	Table ("add", "layer", sLayer, "Notes in Green", "current", Table("find", "color", "Green"));

   SetSymbolLibrary( sPathJOBST + "\\JOBST.SLB") ;   	     
   Symbol( "find", "TextAsSymbol") ;
   hTextEnt = AddEntity ("symbol", "TextAsSymbol", xyHorizStart.x + (xyHorizEnd.x- xyHorizStart.x)/2, xyHorizStart.y + .125, 1, 1, 0) ;  

// Create ID string
   sZipperID =  sStyleID + MakeString("scalar", UID("get", hTextEnt )) ;
   SetDBData(hTextEnt, "ID", sZipperID );
   SetDBData(hTextEnt, "Zipper", "1" );
   SetDBData(hTextEnt, "Data", "INVALID ZIPPER CALCULATION" ); //Do this just incase it fails

   if (DrawCurve) {
       StartPoly("fitted");
	AddVertex ( xyCurveStart);
   	aAngle = Calc("angle" ,xyConstruct, xyHorizStart ) ;
   	aPrevAngle =  Calc("angle" ,xyConstruct, xyCurveStart ) ;
   	if (aAngle > aPrevAngle ) aAngleInc = (aAngle - aPrevAngle) / 4 ;
		else aAngleInc = ((aAngle + 360) - aPrevAngle) / 4 ;	
   	ii = 1 ;
  	while ( ii  <= 3) {	//*
		xyTmp = CalcXY ("relpolar", xyConstruct , nHeelRad, aPrevAngle + aAngleInc * ii); //*
		if (xyTmp.x < xyAnkle.x) {
			AddVertex ( xyTmp); //*
			xyEnd = xyTmp;
			}
		ii = ii + 1 ;
		}
	AddVertex ( xyHorizStart ) ;
	AddVertex ( xyHorizStart.x + 0.25, xyHorizStart.y) ;
	AddVertex ( xyHorizStart.x + 0.5, xyHorizStart.y) ;		// Note this 0.5 is added below for nZipLength > 0
       EndPoly();	
       
       	hEnt = UID("find", UID("getmax")) ;	
       	PR_SetZipperID () ;
       	SetDBData (hEnt, "ZipperLength", "%Length");
        	GetDBValue (hEnt, "ZipperLength", &sTmp, &nDrawnLength);
	nVertex = GetVertexCount(hEnt) + 1;
	
	// Requested zipper length
	if (nZipLength > 0 ) {
		xyHorizEnd.x = xyHorizStart.x + 0.5 +  ((nZipLength / 1.2 ) - nDrawnLength) ;
		xyHorizEnd.y = xyHorizStart.y ;
		if (xyHorizEnd.x > xyFold.x){
			//Draw the bit in the chap
			nZipperInChap = xyHorizEnd.x - xyFold.x ;
			xyZipperStartInChap.x = xyFold.x ;
  	        	xyZipperStartInChap.y = xyFold.y - 1.125;
  	        	xyZipperEndInChap.x = xyHorizEnd.x ;
   		       	xyZipperEndInChap.y = xyFold.y - 1.125;
			xyHorizEnd.x = xyFold.x ;
			nElastic = 0 ;
			}
		SetVertex ( hEnt, nVertex , xyHorizEnd ) ;
		}
	else
		{
		SetVertex ( hEnt, nVertex , xyHorizEnd ) ;
       		GetDBValue (hEnt, "ZipperLength", &sTmp, &nDrawnLength);
		nZipLength = (nZipperInChap + nDrawnLength + nElastic ) * 1.2 ;
		}
	
       	hEnt =AddEntity("marker","closed arrow", xyCurveStart , 0.5 ,0.125, aPrevAngle + 92 ) ;	
       	PR_SetZipperID() ;
       	}
else	{
	if (nZipLength > 0 ) {
		xyHorizEnd.x = xyHorizStart.x + ((nZipLength / 1.2 ))  ;
		xyHorizEnd.y = xyHorizStart.y ;
		}
	else	{
   		nDrawnLength = Calc("length",xyHorizStart, xyHorizEnd) ;
		nZipLength = (nDrawnLength + nElastic ) * 1.2 ;
		}
   	hEnt = AddEntity ("line", xyHorizStart, xyHorizEnd) ;
	PR_SetZipperID() ;
  	hEnt = AddEntity("marker","closed arrow", xyHorizStart , 0.5 ,0.125, 0 ) ;	
	PR_SetZipperID() ;
   	}
   	 	
// Add label and arrows
   
// If a zipper extends to the EOS then we draw this bit
   if (nZipperInChap != 0) {
   	hEnt = AddEntity ("line", xyZipperStartInChap, xyZipperEndInChap) ;
	PR_SetZipperID() ;
//	hEnt = AddEntity("marker","closed arrow",xyZipperStartInChap , 0.5 ,0.125, 0 ) ;	
//  	PR_SetZipperID (); 
	hEnt = AddEntity("marker","closed arrow",xyZipperEndInChap , 0.5 ,0.125, 180 ) ;	
  	PR_SetZipperID (); 
  	  	// Get Velcro overlap (If it exists!) and delete it.
   	sTmp = "DB ID ='" + sStyleID   + "' AND DB Zipper ='VelcroOverlap'";
 	hChan = Open ("selection", sTmp );
   	if (hChan) {
		ResetSelection(hChan);
		while (hEnt = GetNextSelection(hChan)) {
			GetEntityData(hEnt, "layer", &hLayer) ;
			GetEntityClass(hEnt, &sClass);
			if ( hLayer == hLayerNotes && StringCompare("line", sClass) ) 
				hEntClosure = hEnt ;
			else	
				DeleteEntity(hEnt) ;
			}
   		}
	Close("selection", hChan); 
	if (hEntClosure) {
		hLayer = Table("find", "layer", "Template" + sLeg);
		if (hLayer != %badtable) 
			{
			SetEntityData (hEntClosure, "layer", hLayer ) ;
			DeleteDBData (hEntClosure, "Zipper") ;
			}
		}

   	}  
   else
   	{
   	hEnt = AddEntity("marker","closed arrow",xyHorizEnd , 0.5 ,0.125, 180 ) ;	
   	PR_SetZipperID (); 
   	}
   
// Update text symbol with Zipper Length
   if (nMedial == 1 )
   	sTmp = Format("length",nZipLength) +  " MEDIAL ZIPPER" ;
   else
      	sTmp =  Format("length",nZipLength) +  " LATERAL ZIPPER" ;
   SetDBData(hTextEnt, "Data", sTmp );
   

// Reset and exit
   Execute ("menu", "SetLayer",  Table("find", "layer", "1")) ;
   Exit (%ok, "Zipper drawing Complete");

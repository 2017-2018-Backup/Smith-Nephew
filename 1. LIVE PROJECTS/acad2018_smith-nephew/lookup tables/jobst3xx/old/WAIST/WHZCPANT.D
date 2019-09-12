// File Name:	WHZCPANT.D
// Purpose:	Draws a Zip from DISTAL EOS up
//		CHAP style Leg (panty)
// Version:	1.01
// Date:		27.Nov.1996
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action

HANDLE
	hDlg,
	hEnt,
	hChan,
	hCurve,
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
	xyMax, xyMin,
	xyAnkle,
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
	sNoteID,
	sElasticDistal,
	sElasticProximal,
	sElastic,
	sDlgProximalElasticList, 
        	sDlgDistalElasticList ,
	sElasticNote,
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
	nElasticProximal,
	nType,
	ii, nn,
	nZipLength;

BOOL
	EOS,
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
procedure PR_SetZipperID () {
	// Label entity with ID string and make Zipper %true
   	SetDBData(hEnt, "ID", sZipperID );
   	SetDBData(hEnt, "Zipper", "1" );
}

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
   sStyleID         = sStyleID + sLeg ;
   sOtemplateID = sStyleID + "Origin";
   sFoldID	          = sStyleID  + "Fold";
   sAnkleID         = sStyleID +  "Ankle";
   sEOSID	          = sStyleID  + "EOS";
   sNoteID	  = sStyleID  + "PantyElasticNote";
   sTmp = "(type = 'symbol' AND DB  ID ='"  + sNoteID  + "' ) OR (type = 'Marker' AND (DB ID ='"  + sAnkleID  + "'OR DB ID ='"  + sFoldID  + "'OR DB ID ='"  + sEOSID  + "' OR DB ID ='" + sOtemplateID + "'))";
   nMarkersFound = 0 ;
   nMarkersRequired = 3 ;
   Panty = %true;
   hChan = Open ("selection", sTmp );
   if (hChan) {
	ResetSelection(hChan);
	while (hEnt = GetNextSelection(hChan)) {
		GetDBData(hEnt,"ID", &sTmp) ;
		GetGeometry(hEnt, &sTmp1, &xyPt1, &xyTmp, &aAngle) ;
	            if ( xyPt1.x != 0 || xyPt1.y != 0) {
			if (StringCompare(sTmp, sOtemplateID)) {
				nMarkersFound = nMarkersFound + 1 ;
				xyOtemplate = xyPt1 ;
				GetDBData(hEnt, "units", &sUnits);
				GetDBData(hEnt, "Pressure", &sTemplate);
   				GetDBData(hEnt, "Data", &sType); 	
				}
			if (StringCompare(sTmp, sFoldID )){
				nMarkersFound = nMarkersFound + 1 ;
				xyFold = xyPt1 ;
				}
			if (StringCompare(sTmp, sNoteID )){
				GetDBData(hEnt, "Data", &sElasticNote);
				}
			if (StringCompare(sTmp, sAnkleID)) {
				xyAnkle = xyPt1 ;
				Panty = %false ;
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
    if (nMarkersFound < nMarkersRequired ) Exit (%cancel, "Missing markers for selected leg profile, data not found!\n") ;
    if (nMarkersFound > nMarkersRequired ) Exit (%cancel, "Two or more drawings of the same style exist!\nDelete the extra drawing/s and try again.\n") ;
    if (!Panty) Exit (%cancel, "This style has a foot so it is not a Panty!.\n") ;

// Get Smallest circumference by looping through the vetex of the curve 
//
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
// Get Zipper style
//
   sUnits = "Inches";
   EOS = %false;
   Loop = %true ;
   Select = %false ;
   sDlgElasticList =  "3/8\" Elastic\n3/4\" Elastic\n1½\" Elastic\nNo Elastic";
   sDlgProximalElasticList =  "No Elastic\n" + sDlgElasticList;
   if (StringCompare(sElasticNote, "NO ELASTIC"))
   	sDlgDistalElasticList =  "No Elastic\n" + sDlgElasticList;
   else
   	sDlgDistalElasticList =  "3/4\" Elastic\n" + sDlgElasticList;
   sDlgLengthList =  "EOS\nGive a length\nSelected Point\n";
   nMedial = 0 ;
   
   while (Loop) {		
            nButX = 70; nButY = 65;
            hDlg = Open ("dialog", sLeg + " Distal EOS Zipper (CHAP)", "font Helv 8", 20,20, 235, 95);

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
   
    if (Select || nZipLength > 0 ) //closed zipper at proximal end
     		nElasticProximal = 0 ;
    else {
     		nElasticProximal = 0.75 ;
     		if (StringCompare(StringMiddle(sElasticProximal,1,1), "N")) nElasticProximal = 0 ;
     		if (StringCompare(StringMiddle(sElasticProximal,1,3), "3/8")) nElasticProximal = 0.375 ;
     		if (StringCompare(StringMiddle(sElasticProximal,1,1), "1")) nElasticProximal = 1.5 ;
		}
    nElasticDistal = 0.75 ;
    if (StringCompare(StringMiddle(sElasticDistal,1,1), "N")) nElasticDistal = 0 ;
    if (StringCompare(StringMiddle(sElasticDistal,1,3), "3/8")) nElasticDistal = 0.375 ;
    if (StringCompare(StringMiddle(sElasticDistal,1,1), "1")) nElasticDistal = 1.5 ;

    nElastic = nElasticDistal + nElasticProximal ;

    nZipOff = 1.125 ;
    GetVertex(hCurve,1, &xyHorizStart) ;
    xyHorizStart.y = xySmallestCir.y - nZipOff  ;	
    xyHorizEnd.y = xyHorizStart.y ;
	
 //if  EOS given
   if (EOS) {
          if (nMedial == 1 ) {
	 	 xyHorizEnd.x = xyFold.x ;  
    		 nElastic = nElasticDistal ;
   		 nZipperInChap = 0 ;
		}
          else{
          	xyZipperStartInChap.x = xyFold.x ;
          	xyZipperStartInChap.y = xyFold.y - 1.125 ;
          	xyZipperEndInChap = xyEOS ;
          	xyZipperEndInChap.y = xyFold.y - 1.125;
          	nZipperInChap = Calc("length", xyZipperStartInChap, xyZipperEndInChap) ;        	
          	xyHorizEnd.x = xyFold.x ;    
		}  
	  xyHorizEnd.y = xyHorizStart.y ;      
          } 
   else {
   	// No elastic at proximal as this implies a closed zipper
    	nElastic = nElasticDistal  ;
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
   	 	
// Add label and arrows
// If a zipper extends to the EOS then we draw this bit
   if (nZipperInChap != 0) {
   	hEnt = AddEntity ("line", xyZipperStartInChap, xyZipperEndInChap) ;
	PR_SetZipperID() ;
//	hEnt = AddEntity("marker","closed arrow",xyZipperStartInChap, 0.5 ,0.125, 0 ) ;	
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
   else	{
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

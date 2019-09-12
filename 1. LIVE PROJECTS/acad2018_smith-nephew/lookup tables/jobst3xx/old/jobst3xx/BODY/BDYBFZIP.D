// File Name:	BDYBRZIP.D
// Purpose:	Draws a BodyBrief Zipper
//		Waist Height
// Version:	1.01
// Date:		07.07.1997
// Author:		Gary George
//		Et. Al.
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action

HANDLE
	hDlg,
	hEnt,
	hBack,
	hFront,
	hCurve,
	hTextEnt ,
	hLayer;

XY
	xyPt1,
	xyPt2,
	xyPt3,
	xyText,
	xyBackB,		// Back at Buttocks
	xyBackW,	// Back at Waist
	xyBackN, 	// Back at Neck
	xyFrontB,		// Front at Buttocks
	xyFrontW	,	// Front at Waist
	xyFrontC,		// Front at Chest
	xyFrontN,		// Front at Neck
	xySnap,		// Snap crotch 	
	xyPt1,
	xyTmp;

STRING
	sClass,
	sZipperID ,
	sElasticProximal,
	sZipType,
   	sDlgLengthList ,
   	sDlgZipTypeList,
	sUnits,
	sZipLength,
	sProfileID,
	sTmp, sTmp1,
	sLeg,
	sStyleID,
	sError,
	sType;

NUMBER
	ii, nn,
	
	nZipLength ;
BOOL
	bSnapCrotch,
	bBack,
   	bEOStoSelectedPoint,
  	bEOStoStandard,
	bFrontZipper,
	bBackZipper,
	bLoop ;
MSGID
	Ok;
	
// Find JOBST installed directory
// Set path to macros N.B. RE-USE OF CODE
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
   if(!hCurve = GetEntity ("Select a Body Profile", "all")) Exit (%ok, "User cancelled") ;

   Display ("cursor", "wait", "Searching for more data");
   GetDBData (hCurve, "ID", &sStyleID );

// Get Marker data
// Revise style ID to find body bits
//   
   sTmp = "(DB Zipper ='Back*' OR DB Zipper = 'Front*' OR DB Zipper = 'Snap') AND DB ID = '" + sStyleID +  "' " ;
   nEntitiesFound = 0 ;
   nEntitiesRequired = 9 ;
   bSnapCrotch = %false ;
   hChan = Open ("selection", sTmp );
   if (hChan) {
	ResetSelection(hChan);
	while (hEnt = GetNextSelection(hChan)) {
		GetDBData(hEnt,"Zipper", &sTmp) ;
		if (StringCompare(sTmp, "Front" )) {
			nEntitiesFound = nEntitiesFound + 1 ;
			hFront = hEnt ;
			} 
		if (StringCompare(sTmp, "Back" )) {
			nEntitiesFound = nEntitiesFound + 1 ;
			hBack = hEnt ;
			} 
		if (StringCompare(sTmp, "FrontB")) {
			nEntitiesFound = nEntitiesFound + 1 ;
			GetGeometry(hEnt, &sTmp, &xyFrontB , &xyTmp, &aAngle) ;
			}
		if (StringCompare(sTmp, "FrontW")) {
			nEntitiesFound = nEntitiesFound + 1 ;
			GetGeometry(hEnt, &sTmp, &xyFrontW , &xyTmp, &aAngle) ;
			}
		if (StringCompare(sTmp, "FrontC")) {
			nEntitiesFound = nEntitiesFound + 1 ;
			GetGeometry(hEnt, &sTmp, &xyFrontC , &xyTmp, &aAngle) ;
			}
		if (StringCompare(sTmp, "FrontN")) {
			nEntitiesFound = nEntitiesFound + 1 ;
			GetGeometry(hEnt, &sTmp, &xyFrontN , &xyTmp, &aAngle) ;
			}
		if (StringCompare(sTmp, "BackB")) {
			nEntitiesFound = nEntitiesFound + 1 ;
			GetGeometry(hEnt, &sTmp, &xyBackB , &xyTmp, &aAngle) ;
			}
		if (StringCompare(sTmp, "BackW")) {
			nEntitiesFound = nEntitiesFound + 1 ;
			GetGeometry(hEnt, &sTmp, &xyBackW , &xyTmp, &aAngle) ;
			}
		if (StringCompare(sTmp, "BackN")) {
			nEntitiesFound = nEntitiesFound + 1 ;
			GetGeometry(hEnt, &sTmp, &xyBackN , &xyTmp, &aAngle) ;
			}
		if (StringCompare(sTmp, "Snap")) {
			nEntitiesFound = nEntitiesFound + 1 ;
			bSnapCrotch =%true;
			GetGeometry(hEnt, &sTmp, &xySnap , &xyTmp, &aAngle) ;
			}
		}
   	}
   Close("selection", hChan); 

// Check if that the ENTITIES have been found, otherwise exit
//
  if (nEntitiesFound  < nEntitiesRequired ) Exit (%cancel, "Missing marker for selected Body Suit, data not found!\n") ;
  if (nEntitiesFound  > nEntitiesRequired ) Exit (%cancel, "Two or more drawings of the same style exist!\nDelete the extra drawing/s and try again.\n") ;

// Create Dialog
// Get Zipper style
//

   bEOStoSelectedPoint = %false ;
   bEOStoStandard = %false ;
   sDlgLengthList =  "Standard\nGive a length\nSelected Point";
   sDlgZipTypeList =  "\nBack Zipper\nFront Zipper";
   Loop = %true ;
   nMinZipLength = 5 ;
   sError = "" ;
   while (Loop) {		
            nButX = 65; nButY = 40;
            hDlg = Open ("dialog", sLeg + " Body Zipper (Body Brief)", "font Helv 8", 20,20, 220, 75);

	AddControl ( hDlg, "pushbutton",  nButX, nButY, 35, 14, "Cancel", "%cancel", "");
	AddControl ( hDlg, "pushbutton",  nButX + 48 , nButY, 35,14, "OK", "%ok", "");

	AddControl ( hDlg, "ltext", 5, 12, 28, 14 , "EOS to", "string", "");
	AddControl ( hDlg, "combobox", 30, 10, 70, 40 , sDlgLengthList , "string", "sZipLength");
	AddControl ( hDlg, "ltext", 105, 12, 50, 20 , "Zipper Type:", "string", "");
 	AddControl ( hDlg, "dropdownlist", 150, 10, 60, 70 , sDlgZipTypeList , "string", "sZipType");
	AddControl ( hDlg, "ltext", 5, nButY + 25, 200, 14 , sError, "string", "");

      	Ok = Display("dialog", hDlg, "%center");
  	Close("dialog", hDlg);

	if (Ok == %cancel ) Exit (%ok, "User Cancel!") ;	
	
	if (StringCompare("Selected Point", sZipLength)) bEOStoSelectedPoint = %true;
	if (StringCompare("Standard", sZipLength))  bEOStoStandard = %true ;
	if (StringCompare("Front Zipper", sZipType)) bFrontZipper = %true;
	if (StringCompare("Back Zipper", sZipType)) bBackZipper = %true;

	if (bEOStoStandard)
		Loop = %false ;
	else {
		sError = "Error:- Not yet implemented!" ;
		Display ("beep",sError) ;
		Loop = %true ;
		
//		nZipLength = Value("scalar", sZipLength);
//		if ( nZipLength == 0  && StringLength (sZipLength) > 0) {
//			Display ("beep", "Invalid given length!") ;
//			Loop = %true ;
//			}
//   		else
//			Loop = %false ;
		}
		
	if ( !bFrontZipper && !bBackZipper){
		sError = "Error:- You must select either Front or Back Zipper!" ;
		Display ("beep",sError) ;
		Loop = %true ;
		}
    }   
    
   Display ("cursor", "wait", "Drawing Zipper");
 
// Calculate zipper text position
// 
   if (bBackZipper){
	GetDBValue (hBack, "ZipperLength", &sTmp, &nZipLength );
	xyText.x =xyBackW.x + (xyBackN.x- xyBackW.x)/2;
	xyText.y = xyBackW.y + .125;
	nLength = Calc("length", xyBackB, xyBackW) ;
	aAngle =  Calc("angle", xyBackB, xyBackW); 
	xyPt1 = CalcXY ("relpolar", xyBackB, nLength/2, aAngle ) ;
	xyPt2 = xyPt1;
	xyPt2.y = xyPt1.y + 0.125;
	xyPt3 = CalcXY ("relpolar", xyPt1, 0.5, aAngle) ;	
   	}
// Draw on layer Notes
//
   hLayer = Table("find", "layer", "Notes") ;
   if ( hLayer > %zero )
 	Execute ("menu", "SetLayer",  hLayer) ;
   else
   	Table ("add", "layer", sLayer, "Notes in Green", "current", Table("find", "color", "Green"));

// 
// Draw Zipper Text firsts (GET ID FOR LATER USE)
// 
   SetSymbolLibrary( sPathJOBST + "\\JOBST.SLB") ;   	     
   Symbol( "find", "TextAsSymbol") ;
   hTextEnt = AddEntity ("symbol", "TextAsSymbol", xyText , 1, 1, 0) ;  

// Create ID string
   sZipperID =  sStyleID + MakeString("scalar", UID("get", hTextEnt )) ;
   SetDBData(hTextEnt, "ID", sZipperID );
   SetDBData(hTextEnt, "Zipper", "1" );
   sZipLength = Format("length", nZipLength);
   SetDBData(hTextEnt, "Data", sZipLength+" Back Zipper" );
       	  
// Draw mark
   hEnt = AddEntity ("line", xyPt1, xyPt2) ;
   PR_SetZipperID() ;
   hEnt = AddEntity ("line", xyPt2, xyPt3) ;
   PR_SetZipperID() ;
	

// Reset and exit
   Execute ("menu", "SetLayer",  Table("find", "layer", "1")) ;
   Exit (%ok, "Zipper drawing Complete");

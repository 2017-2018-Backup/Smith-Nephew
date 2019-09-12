// File Name:	ZIPULNAR.D
// Purpose:	Draws an ULNAR Zipper (Little Finger Side) for
//		the CAD Glove
// Version:	1.01
// Date:		22.Sep.95
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action

   Display ("cursor", "wait", "Loading L.F.S. Zipper Tool");

HANDLE
	hDlg,
	hEnt,
	hEOS,
	hLFA,
	hLFS,	
	hChan,
	hCurve,
	hLayer;

XY
	xyZipStart,
	xyZipEnd,
	xyPALMER,
	xyDORSAL,
	xyPALMERWEB,
	xyStartEOS, xyEndEOS,
	xyGivenPoint,
	xyPt1,
	xyPt2,
	xyCircleCen, 
	xyTmp;

STRING
	sClass,
	sText,
	sAge,
	sElasticProximal,
	sUnits,
	sZipLength,
	sZipOffset,
	sProfileID,
	sTmp, sTmp1,
	sSide,
	sDlgElasticList ,
	sDlgLengthList,
	sDlgWebOffSetList ,
	sWebOffSet,
	sZipperID; // Unique ID for each Zipper, allows easy selection and deletion

NUMBER
	nFound,
	nRadius,
	nAge,
	nRequired,
	nInsertSize,
	nElasticFactor,
	nElasticProximal,
	nElastic,	//sum of Proximal and Distal
	ii, nn,
	nLenLFS,
	nWebOffSet,
	nZipLength ;
BOOL
	EOStoSelectedPoint,
	EOStoGivenLength,
	EOStoCalculatedPoint,
	Loop ;
MSGID
	Ok;

ANGLE
	aPALMER,
	aPALMERWEB,
	aEOS,
	aZipper,
	aDelta,
	aDeltaLFA,
	aAngle;
	

// Clear user selection
//
   UserSelection ("clear") ;
   UserSelection ("update") ;

   STRING		sPathJOBST;
   sPathJOBST = GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;
   SetData("PathMacro",  sPathJOBST + "\\GLOVECAD" );
   
// Load Functions
//
   @ZIPDEF.D;
   procedure GetIntersection() {
      	xyPt1 = CalcXY ("relpolar", xyPALMERWEB,  nWebOffSet + nInsertSize , (aZipper + (nDirection*180))) ;   	
   	xyPt2 = CalcXY ("relpolar", xyPt1, 2, (aEOS + (nDirection*180))) ;  
	if (FN_CirLinInt( xyPt2, xyPt1, xyCircleCen, nRadius)) 
		xyZipEnd = xyInt;
	else 
		Exit(%abort,  "Unable to find end point of Zipper." ) ;	
   } // End GetIntersection

   
// Get profile, identify side 
//
   if (! hCurve=GetEntity ("Select Glove Profile", "all")) Exit (%ok, "User Cancelled") ;

   Display ("cursor", "wait", "Searching for more data");
   GetDBData (hCurve, "ID", &sProfileID);
   nStringLength = StringLength(sProfileID);
   if (StringCompare( "Left", StringMiddle(sProfileID, nStringLength - 3, 4))){
	sSide = "Left";
	nDirection = 1;
	}
    if (StringCompare( "Right", StringMiddle(sProfileID, nStringLength - 4, 5))) {
	sSide = "Right";
	nDirection = -1 ;
	}

// Check if this is a valid Glove exit if not
// This is crude check only
    if (StringLength (sSide) == 0) Exit (%cancel, "A Glove Profile was not selected\n") ;

// Get data for an ULNAR Zipper 
   sTmp = "DB ID = '" + sProfileID + "' AND  (DB Zipper ='PALMER*' OR DB Zipper ='LFA' OR DB Zipper ='LFS'  OR DB Zipper = 'EOS')";
   nFound = 0 ; nReqired = 5;
   hChan = Open ("selection", sTmp );
   if (hChan) {
	ResetSelection(hChan);
	while (hEnt = GetNextSelection(hChan)) {
		sTmp = "" ;
		GetDBData(hEnt,"Zipper", &sTmp) ;
		if (StringCompare(sTmp, "PALMER")) {
			nFound = nFound + 1 ;
			GetGeometry(hEnt, &sTmp, &xyPALMER , &xyTmp, &aAngle) ;
			GetDBData(hEnt, "Age", &sAge);
			nAge = Value("scalar", sAge) ;
			}
		if (StringCompare(sTmp, "PALMER-WEB" )){
			nFound = nFound + 1 ;
			GetGeometry(hEnt, &sTmp, &xyPALMERWEB, &xyTmp, &aAngle) ;
			// We need slant insert size for ULNAR Zippers
			GetDBValue(hEnt, "ZipperLength", &sTmp1, &nInsertSize);
			}
		if (StringCompare(sTmp, "EOS" )){
			// End Of Support 
			nFound = nFound + 1 ;
			hEOS = hEnt ;
			GetGeometry(hEnt, &xyStartEOS, &xyEndEOS) ;
			}
		if (StringCompare(sTmp, "LFA" )){
			// Little Finger Arc
			nFound = nFound + 1 ;
			hLFA = hEnt ;
   			GetGeometry (hLFA, &xyCircleCen, &nRadius, &aAngle, &aDeltaLFA);
			}
		if (StringCompare(sTmp, "LFS" )){
			//  Little Finger Straight (line from EOS to start of LFA above)
			nFound = nFound + 1 ;
			hLFS = hEnt ;
			GetDBValue(hEnt, "ZipperLength", &sTmp1, &nLenLFS);
			}
		}
   	}
   Close("selection", hChan); 

// Check that sufficent data have been found, otherwise exit
//
    if (!hLFA) Exit (%cancel, "Can't draw a zip with this module!\nTry using the LFS Zip to Elbow") ;
    if (nFound < nReqired) Exit (%cancel, "Missing data for selected Glove!\n") ;
    if (nFound > nReqired) Exit (%cancel, "Two or more drawings of the same style exist!\nDelete the extra drawing/s and try again.\n") ;

// Create Dialog
// Get Zipper style
//
   EOStoSelectedPoint = %false ;
   EOStoCalculatedPoint = %false ;
   EOStoGivenLength = %false ;  
   
// Proximal elastic
   sDlgElasticList =  "3/8\" Elastic\n3/4\" Elastic\n1½\" Elastic\nNo Elastic";
   if (nAge < 10)
   	sDlgElasticList =  "3/8\" Elastic\n" + sDlgElasticList;  // 1/2" for children under 10
   else
      	sDlgElasticList =  "3/4\" Elastic\n" + sDlgElasticList;  // Inch for adults 

// Length specification      	
   sDlgLengthList =  "Standard\nGive a length\nSelected Point";
   
// Offset from web
   sDlgWebOffSetList =  "1-1/8\"\n3/4\"";
   sDlgWebOffSetList =  "1-1/8\"\n"  + sDlgWebOffSetList ;	// Set Default

   Loop = %true ;
   while (Loop) {		
            nButX = 65; nButY = 55;
            hDlg = Open ("dialog", sSide + " Glove Zipper (L.F.S.)", "font Helv 8", 20,20, 210, 75);

	AddControl ( hDlg, "pushbutton",  nButX, nButY, 35, 14, "Cancel", "%cancel", "");
	AddControl ( hDlg, "pushbutton",  nButX + 48 , nButY, 35,14, "OK", "%ok", "");

	AddControl ( hDlg, "ltext", 5, 12, 28, 14 , "EOS to", "string", "");
	AddControl ( hDlg, "combobox", 35, 10, 70, 40 , sDlgLengthList , "string", "sZipLength");

	AddControl ( hDlg, "ltext", 110, 12, 30, 14 , "Proximal:", "string", "");
 	AddControl ( hDlg, "dropdownlist", 145, 10, 60, 70 , sDlgElasticList, "string", "sElasticProximal");

	AddControl ( hDlg, "ltext", 100, 32, 50, 14 , "Below Web:", "string", "");
 	AddControl ( hDlg, "dropdownlist", 145, 30, 60, 40 , sDlgWebOffSetList, "string", "sWebOffSet");

      	Ok = Display("dialog", hDlg, "%center");
  	Close("dialog", hDlg);

	if (Ok == %cancel ) Exit (%ok, "User Cancel!") ;	
	
	if (StringCompare("Selected Point", sZipLength)) EOStoSelectedPoint = %true;
	
	if (StringCompare("Standard", sZipLength)) EOStoCalculatedPoint = %true;
	    
	if (EOStoSelectedPoint || EOStoCalculatedPoint   )
		Loop = %false ;
	else {
		nZipLength = Value("scalar", sZipLength);
		if ( nZipLength == 0  && StringLength (sZipLength) > 0) {
			Display ("message", "error", "Invalid given length!\n\nTo use this option, type over the text in the \"EOS to:\" box with the required length in Inches (Decimal Inches).") ;
			Loop = %true ;
			}
   		else {
			Loop = %false ;
			EOStoGivenLength = %true ;  
			}
		}
    }   
    
   Display ("cursor", "wait", "Drawing Zipper");
	
    	
//Establish allowance for zippers
//
   nElasticProximal = 0.75 ;
   if (StringCompare(StringMiddle(sElasticProximal,1,1), "N")) nElasticProximal = 0 ;
   if (StringCompare(StringMiddle(sElasticProximal,1,3), "3/8")) nElasticProximal = 0.375 ;
   if (StringCompare(StringMiddle(sElasticProximal,1,1), "1")) nElasticProximal = 1.5 ;
   nElastic = nElasticProximal ;
   
// Elastic factor   
// Set by age
   if (nAge <= 6 )
   	nElasticFactor = 1;
   else	
   	nElasticFactor = 0.95;
    
// Establish Minimum offset from web    
   nWebOffSet = 1.125 ;
   if (StringCompare(StringMiddle(sWebOffSet,1,1), "3")) nWebOffSet = 0.75 ;
   
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
// Establish EOS start and end
//
   nLengthToStart = Calc ("length", xyStartEOS, xyPALMER);	
   nLengthToEnd = Calc ("length", xyEndEOS, xyPALMER);
   if (nLengthToStart  > nLengthToEnd){
   	// Swap start and end
   	xyTmp = xyStartEOS ;
   	xyStartEOS =  xyEndEOS ;
   	xyEndEOS = xyTmp ; 
   	}
   xyZipStart = xyStartEOS;

// Angles, Construction lengths etc
//
   aEOS = Calc ("angle", xyStartEOS, xyEndEOS );        
   aPALMERWEB = Calc ("angle", xyStartEOS, xyPALMERWEB );
   aAngle = abs(aEOS - aPALMERWEB) ;

// Perpendicular Distance from EOS to xyPALMERWEB 
   nLength = Calc ("length", xyStartEOS, xyPALMERWEB );
   nZipConstructLen = abs(sin(aAngle) * nLength);

// Zipper angle     
   aZipper = aEOS + (90 * nDirection);
   
   if (EOStoCalculatedPoint) GetIntersection();
     	 	
   if (EOStoSelectedPoint){
   	// Only allow the point if it lies within the rectangle formed by the
   	// EOS and the xyPALMERWEB plus a 1/16th" tolerance
   	// EOS = length and angle, xyPALMERWEB = width
    	xyPt1 = CalcXY ("relpolar", xyStartEOS, nZipConstructLen + .0625, aZipper) ; 
   	xyPt1 = CalcXY ("relpolar", xyPt1,  2, (aEOS + (nDirection*180))) ;   	
   	xyPt2 = CalcXY ("relpolar", xyEndEOS, nZipConstructLen + .0625, aZipper) ; 
   	xyPt2 = CalcXY ("relpolar", xyPt2,  2, aEOS ) ;   	

     	Loop = %true;
   	while(Loop ){
  		if (!GetUser("xy", "Select End of Zipper", &xyGivenPoint)) Exit(%ok, "Start not selected") ;
		if ( !Calc ("inpoly", xyGivenPoint, xyStartEOS, xyPt1, xyPt2, xyEndEOS) ) {
			Ok = Display("message", "OKquestion", "Given point can't be used! Try again.\nOr use cancel to Exit");
			if ( Ok == %cancel) Exit (%ok, "User Cancelled");
			}
		else
			Loop = %false;
		}	
		
   	xyPt1 = CalcXY ("relpolar", xyGivenPoint ,  5, aEOS ) ;   	
   	xyPt2 = CalcXY ("relpolar", xyPt1, 7, (aEOS + (nDirection*180))) ;   	
	if ( FN_CirLinInt( xyPt1, xyPt2, xyCircleCen, nRadius)) 
		xyZipEnd = xyInt;
	else 
		Exit(%abort,  "Unable to find end point of Zipper." ) ;		
		
   	aStart = Calc ("angle", xyStartEOS, xyGivenPoint);
   	aAngle = abs( aEOS - aStart) ;
  	nLength = Calc ("length", xyStartEOS, xyGivenPoint );
  	nLength = abs(sin(aAngle) * nLength);

      	// Ensure that selected point is no closer than nWebOffSet to the web
   	// within 1/8th of an inch
   	nLength = nLength - nZipConstructLen   ;
	if ( abs(nLength) < (nWebOffSet - 0.125) ) {
		sText = "The end of the zipper will be closer to the nearest web\\slant insert than " 
			+ Format ("length",nWebOffSet)
			+ ".  Actual distance is " 
			+ Format ("length",abs(nLength)) ;
		Ok = Display ("message", "YNCquestion", sText  + "\nUse YES to use this point or\nUse NO to default to Standard." ) ;
		if (Ok == %cancel) Exit (%ok, "User Cancelled");
		if (Ok == %no) {
			if ( FN_CirLinInt( xyPt1, xyPt2, xyCircleCen, nRadius)) 
				xyZipEnd = xyInt;
			else 
				Exit(%abort,  "Unable to find end point of Zipper." ) ;		
			}
		}
	}


   if  (EOStoGivenLength ) {
      	nArcLength =((nZipLength * nElasticFactor ) - nElastic) - nLenLFS ; 
      	aDelta = (nArcLength * 180) / ( nRadius * %pi) ;
      	if (abs(aDelta) < aDeltaLFA) {
      		aStart = Calc("angle", xyCircleCen, xyPALMER);
  		xyZipEnd = CalcXY ("relpolar", xyCircleCen, nRadius, (aStart - (nDirection*aDelta)) mod 360) ;   	
      		aStart = Calc ("angle", xyStartEOS, xyZipEnd);
   		aAngle = abs(aEOS - aStart) ;
  		nLength = Calc ("length", xyStartEOS, xyZipEnd);
  		nLength = abs (sin(aAngle) * nLength);
		nLength = nLength - nZipConstructLen   ;
         		}
         	else	
         		// The given length will not intersect the arc
         		nLength = 0 ; // Dummy to force check below 

	if ( (abs(nLength) < (nWebOffSet - 0.125)) || (nLength > 0)) {
		Ok = Display ("message", "OKquestion", 
			 "The end of the zipper will be above or closer to the nearest web\\slant insert than " 
			+ Format ("length",nWebOffSet) + " with the given length."
			+  "\nUse OK to default to Standard." ) ;
		if (Ok == %cancel) Exit (%ok, "User Cancelled");
		GetIntersection() ;
		nZipLength = 0 ;
		}
	}  
	
// Calculate length of zipper
//
    if ( nZipLength == 0) {
    	aStart = Calc("angle", xyCircleCen, xyPALMER);
    	aAngle = Calc("angle", xyCircleCen, xyZipEnd);
    	aDelta = max (aAngle, aStart) - min (aAngle, aStart)  ;   	
    	if (aDelta > 180) aDelta = 360 - aDelta ;
    	nLength =  (aDelta / 180) * nRadius * %pi ;
   	nDrawnLength = nLenLFS  + nLength; 
    	nZipLength = (nDrawnLength + nElastic ) / nElasticFactor ;
    	}
    sZipLength = Format("length", nZipLength);

// Draw markers and Text label
// Each entity is given an ID Data Base value that links them together
// This is based on the UID of the symbol "ZipperText" and the Glove Profile ID.
// A Symbol is used as DB values can't be attached to text entities 

// Add text symbol
   xyTmp = CalcXY("relpolar", xyPALMER, Calc ("length", xyZipEnd, xyPALMER) / 2, aZipper) ;
   
   SetSymbolLibrary( sPathJOBST + "\\JOBST.SLB") ;   	     
   Symbol( "find", "TextAsSymbol") ;
   
   hEnt = AddEntity ("symbol", "TextAsSymbol", xyTmp, 1, 1, aZipper) ;  
   SetDBData(hEnt, "Data", sZipLength + " L.F.S.Zipper" );
   
// Create ID string
   sZipperID =  sProfileID + MakeString("scalar", UID("get", hEnt)) ;
   
// Label entity with ID string and make Zipper %true
   SetDBData(hEnt, "ID", sZipperID );
   SetDBData(hEnt, "Zipper", "1" );
    	
// Add label and arrows
   hEnt = AddEntity("marker","closed arrow", xyZipStart , 0.5 ,0.125, aEOS + (45 * nDirection )) ;	
   SetDBData(hEnt, "ID", sZipperID );
   SetDBData(hEnt, "Zipper", "1" );
   
   hEnt = AddEntity("marker","closed arrow", xyZipEnd , 0.5 ,0.125, aEOS - (45 * nDirection)) ;		
   SetDBData(hEnt, "ID", sZipperID );
   SetDBData(hEnt, "Zipper", "1" );
   
// Reset and exit
   Execute ("menu", "SetLayer",  Table("find", "layer", "1")) ;
   Exit (%ok, "Zipper drawing Complete");

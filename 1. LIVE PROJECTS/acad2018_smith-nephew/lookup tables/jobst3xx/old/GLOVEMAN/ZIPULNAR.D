// File Name:	ZIPULNAR.D
// Purpose:	Draws an ULNAR Zipper (Little Finger Side) for
//		the CAD Glove
// Version:	1.02
// Date:		28.Sep.95
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
// 11.Apr.96	GG	Revisions w.r.t Manual Glove

   Display ("cursor", "wait", "Loading L.F.S. Zipper Tool");

HANDLE
	hDlg,
	hEnt,
	hEOS,
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
	xyStandardEnd,
	xyText,
	xyTmp, xyTmp1;

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
	sDlgElasticList,
	sDlgLengthList,
	sDlgWebOffSetList ,
	sWebOffSet,
	sData,
	sZipperID; // Unique ID for each Zipper, allows easy selection and deletion

NUMBER
	Ok,
	nType,
	nFound,
	nRadius,
	nAge,
	nReqired,
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

ANGLE
	aPALMER,
	aPALMERWEB,
	aEOS,
	aStart,
	aZipper,
	aDelta,
	aAngle;

// Clear user selection
//
   UserSelection ("clear") ;
   UserSelection ("update") ;
   
   STRING		sPathJOBST;
   sPathJOBST = GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;
   SetData("PathMacro",  sPathJOBST + "\\GLOVEMAN" );
   

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
	nDirection = -1;
	}

// Check if this is a valid Glove exit if not
// This is crude check only
    if (StringLength (sSide) == 0) Exit (%cancel, "A Glove Profile was not selected\n") ;

// Get data for an ULNAR Zipper 
   sTmp = "DB ID = '" + sProfileID + "' AND  (DB Zipper = 'PALMER' OR DB Zipper = 'PALMER-WEB'  OR DB Zipper = 'LFS'  OR DB Zipper = 'EOS')";
   nFound = 0 ; nReqired = 4;
   hChan = Open ("selection", sTmp );
   if (hChan) {
	ResetSelection(hChan);
	while (hEnt = GetNextSelection(hChan)) {
		GetDBData(hEnt,"Zipper", &sTmp) ;
		if (StringCompare(sTmp, "PALMER")) {
			nFound = nFound + 1 ;
			GetGeometry(hEnt, &sTmp, &xyPALMER , &xyTmp, &aAngle) ;
			GetDBData(hEnt, "Age", &sAge);
			GetDBData(hEnt, "Data", &sData);
			nType = Value( "scalar", StringMiddle(sData, 10, 2));
			nAge = Value("scalar", sAge) ;
			}
		if (StringCompare(sTmp, "EOS" )){
			// End Of Support 
			nFound = nFound + 1 ;
			hEOS = hEnt ;
			GetGeometry(hEnt, &xyStartEOS, &xyEndEOS) ;
			}
		if (StringCompare(sTmp, "PALMER-WEB" )){
			// PALMER-WEBweb
			nFound = nFound + 1 ;
			GetGeometry(hEnt, &sTmp, &xyPALMERWEB,  &xyTmp, &aAngle) ;
			GetDBValue(hEnt, "ZipperLength", &sTmp1, &nInsertSize);
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
    if (nFound < nReqired) Exit (%cancel, "Missing data for selected Glove!\n") ;
    if (nFound > nReqired) Exit (%cancel, "Two or more drawings of the same style exist!\nDelete the extra drawing/s and try again.\n") ;
// Load Functions
//
   @ZIPDEF.D;
   
//
//
function FN_PolyInterSection () {        
   nVertex = GetVertexCount (%1) ;
   nn = 1; 
   GetVertex ( %1, nn, &xyTmp1) ; 
   while (nn <= nVertex) {
   	nn = nn + 1;
   	xyTmp = xyTmp1 ;
   	GetVertex ( %1, nn, &xyTmp1) ; 
   	// Extend the first and last segments to 10 inches to ensure intersection
   	if (nn == 2){ xyTmp = CalcXY ("relpolar", xyTmp1, 10,Calc ("angle", xyTmp1, xyTmp));}
   	if (nn == nVertex){ xyTmp1 = CalcXY ("relpolar", xyTmp, 10, Calc ("angle", xyTmp, xyTmp1)); } 		
   	if ( FN_LinLinInt (%2, %3, %4, %5, xyTmp, xyTmp1)) {
   		return ( %true) ;	
		}
  	}
   return(%false) ;	
   } //End FN_PolyInterSection 
   
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
   else	{
   	nElasticFactor = 0.95; 		// Glove to elbow & Normal Glove
   	if (nType == 2) nElasticFactor = 0.92; 	// Glove to axilla
   	}
   	  	
    
// Establish Minimum offset from web    
   nWebOffSet = 1.125 ;
   if (StringCompare(StringMiddle(sWebOffSet,1,1), "3")) nWebOffSet = 0.75 ;
   
// Draw on layer Notes
//
   Execute ("menu", "SetLayer",  Table("find", "layer", "Notes")) ;
   
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
   aAngle = abs (aEOS - aPALMERWEB) ;

// Perpendicular Distance from EOS to xyPALMERWEB 
   nLength = Calc ("length", xyStartEOS, xyPALMERWEB );
   nZipConstructLen = abs(sin(aAngle) * nLength);

// Zipper angle     
   aZipper = aEOS + (90 * nDirection);
    
// Get standard point for later use
//   	
   xyPt1 = CalcXY ("relpolar", xyPALMERWEB,  nWebOffSet + nInsertSize, (aZipper + (nDirection*180))) ;   	
   xyPt1 = CalcXY ("relpolar", xyPt1,  2, (aEOS + (nDirection*180))) ;   	
   xyPt2 = CalcXY ("relpolar", xyPt1, 4, aEOS ) ;   	
	
   if (!FN_PolyInterSection (hLFS, xyPt1, xyPt2)) Exit (%abort, "Can't find end of Zipper intersection on Little Finger Side") ;
//   hEnt = AddEntity("marker","crosshair", xyInt , 0.125 ,0.125, 0) ;
   xyStandardEnd = xyInt ;	// xyInt is a global set by   FN_PolyInterSection
        
   if (EOStoCalculatedPoint) xyZipEnd = xyStandardEnd ;  	 

   if (EOStoSelectedPoint){
   	// Only allow the point if it lies within the rectangle formed by the
   	// EOS and the xyPALMERWEB plus a 1/16th" tolerance
   	// EOS = length and angle, xyPALMERWEB  = width
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
		
		
   	xyPt1 = CalcXY ("relpolar", xyGivenPoint, 10, (aEOS + (nDirection*180))) ; 
   	xyPt2 = CalcXY ("relpolar", xyPt1, 20, aEOS) ;   	
   	if (!FN_PolyInterSection (hLFS, xyPt1, xyPt2)) Exit (%abort, "Can't find end of Zipper intersection on Thumbside") ;
      	xyGivenPoint = xyInt ;	// xyInt is a global set by   FN_PolyInterSection
	xyZipEnd = xyGivenPoint ;
	
      	// Ensure that selected point is no closer than nWebOffSet to the web
   	// within 1/8th of an inch
      	aStart = Calc ("angle", xyStartEOS, xyGivenPoint);
   	aAngle = abs( aEOS - aStart) ;
  	nLength = Calc ("length", xyStartEOS, xyGivenPoint );
  	nLength = abs(sin(aAngle) * nLength);
   	nLength = nLength - (nZipConstructLen - nInsertSize)  ;
	if ( abs(nLength) < (nWebOffSet - 0.125) ) {
		sText = "The end of the zipper will be closer to the nearest web\\slant insert than " 
			+ Format ("length",nWebOffSet)
			+ ".  Actual distance is " 
			+ Format ("length",abs(nLength)) ;
		Ok = Display ("message", "YNCquestion", sText  + "\nUse YES to use this point or\nUse NO to default to Standard." ) ;
		if (Ok == %cancel) Exit (%ok, "User Cancelled");
		if (Ok == %no) xyZipEnd = xyStandardEnd;
		}
	}


   if  (EOStoGivenLength ) {
      	nRequired = (nZipLength * nElasticFactor ) - nElastic ; 
      	nTraversed = 0 ;
      	nVertex = GetVertexCount (hLFS) ;
   	nn = nVertex;
   	GetVertex ( hLFS, nVertex, &xyTmp1) ; 
  	// Account for 3/4" cut back at elbow for glove to elbow  or 1 tape cut back for glove with flap 
   	// 
    	nLength = Calc ("length", xyZipStart, xyTmp1) ;
  	if (nLength > 0.0625) nRequired = nRequired + nLength ;
  	//
    	while (nn >= 1) {
   		nn = nn - 1;
   		xyTmp = xyTmp1 ;
   		GetVertex ( hLFS, nn, &xyTmp1) ; 
   		// Extend the last segments to 100 inches to ensure intersection
	   	if (nn == 1) { xyTmp1 = CalcXY ("relpolar", xyTmp, 10000, Calc ("angle", xyTmp, xyTmp1)); } 		
   		nLength = Calc ("length", xyTmp1, xyTmp) ;
		if ((nTraversed + nLength) >= nRequired ) {
			nLength = nRequired - nTraversed ;
			if (!FN_CirLinInt( xyTmp, xyTmp1, xyTmp, nLength)) Exit (%abort, "Can't establish intersection point") ;
			xyZipEnd = xyInt ;  // xyInt is a global from FN_CirLinInt() above
			break ;
			}
		else
			nTraversed = nTraversed + nLength ;
	  	}

      	// Ensure that selected point is no closer than nWebOffSet to the web
   	// within 1/8th of an inch
      	aStart = Calc ("angle", xyStartEOS, xyZipEnd);
   	aAngle = abs( aEOS - aStart) ;
  	nLength = Calc ("length", xyStartEOS, xyZipEnd);
  	nLength = abs(sin(aAngle) * nLength);
   	nLength = nLength - (nZipConstructLen - nInsertSize)   ;

	if ( (abs(nLength) < (nWebOffSet - 0.125)) || (nLength > 0)) {
		Ok = Display ("message", "OKquestion", 
			 "The end of the zipper will be above or closer to the nearest web\\slant insert than " 
			+ Format ("length",nWebOffSet) + " with the given length."
			+  "\nUse OK to default to Standard." ) ;
		if (Ok == %cancel) Exit (%ok, "User Cancelled");
		xyZipEnd = xyStandardEnd ;  	 
		nZipLength = 0 ;  // Force a calculation of length as below
		}
	}  
   
// Calculate length of zipper
//
    if ( nZipLength == 0) {
      	xyPt1 = CalcXY ("relpolar", xyZipEnd, 2, aEOS) ;   	
  	xyPt2 = CalcXY ("relpolar", xyZipEnd, 2, aEOS + 180 ) ;   	
      	nDrawnLength = 0 ;
      	nVertex = GetVertexCount (hLFS) ;
   	nn = nVertex;
   	GetVertex ( hLFS, nVertex, &xyTmp1) ; 
   	// Account for 3/4" cut back at elbow for glove to elbow  or 1 tape cut back for glove with flap 
   	// 
    	nLength = Calc ("length", xyZipStart, xyTmp1) ;
  	if (nLength > 0.0625) nDrawnLength = nDrawnLength - nLength ;
  	//
   	while (nn >= 1) {
   		nn = nn - 1;
   		xyTmp = xyTmp1 ;
   		GetVertex ( hLFS, nn, &xyTmp1) ; 
   		// Extend the last segment to 100 inches to ensure intersection
	   	if (nn == 1) { xyTmp1 = CalcXY ("relpolar", xyTmp, 100, Calc ("angle", xyTmp, xyTmp1)); } 		
		if (FN_LinLinInt( xyTmp, xyTmp1, xyPt1, xyPt2)) {
			nDrawnLength = nDrawnLength +  Calc ("length", xyTmp, xyInt) ;
			break ;
			}
		else
			nDrawnLength = nDrawnLength +  Calc ("length", xyTmp1, xyTmp) ;
	  	}
    	nZipLength = (nDrawnLength + nElastic ) / nElasticFactor ;
    	}
    sZipLength = Format("length", nZipLength);

// Draw markers and Text label
// Each entity is given an ID Data Base value that links them together
// This is based on the UID of the symbol "ZipperText" and the Glove Profile ID.
// A Symbol is used as DB values can't be attached to text entities 

// Add text symbol
   xyText = CalcXY("relpolar", xyZipStart, Calc ("length", xyZipStart, xyZipEnd) / 2, Calc("angle", xyZipStart, xyZipEnd )) ;

   xyPt1 = CalcXY ("relpolar", xyText ,  10, (aEOS + (nDirection*180))) ;   	
   xyPt2 = CalcXY ("relpolar", xyText , 10, aEOS ) ;   	
	
   if (!FN_PolyInterSection (hLFS, xyPt1, xyPt2)) 
   	Display ("Message", "error", "Can't find a proper position for the Zipper Text. Reposition manually") ;
   else
  	{xyText = CalcXY ("relpolar", xyInt ,  0.25 , aEOS) ;} 

//   hEnt = AddEntity("marker","crosshair", xyInt , 0.125 ,0.125, 0) ;
   	
   
   SetSymbolLibrary( sPathJOBST + "\\JOBST.SLB") ;   	     
   Symbol( "find", "TextAsSymbol") ;
   hEnt = AddEntity ("symbol", "TextAsSymbol", xyText, 1, 1, aZipper) ;  
   SetDBData(hEnt, "Data", sZipLength + " L.F.S." );
   
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

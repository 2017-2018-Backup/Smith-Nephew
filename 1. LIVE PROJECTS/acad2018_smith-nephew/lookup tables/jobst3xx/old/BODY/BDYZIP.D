// File Name:	BDYZIP.D
// Purpose:		Draws a Body Zipper
//		
// Version:	1.0
// Date:		23.07.1997
// Author:		Gary George/Paul O'Rawe
//		
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
	xyBackW,		// Back at Waist
	xyBackN, 		// Back at Neck
	xyFrontB,		// Front at Buttocks
	xyFrontW,		// Front at Waist
	xyFrontC,		// Front at Chest
	xyFrontN,		// Front at Neck
	xySnap,		// Snap crotch 	
	xyRequestedPoint,
	
	xyTmp,xyTmp1;

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
	s1, s2,
	sType;

NUMBER
	ii, nn, 
	nLen,
	nSnapToFrontB,
	nTurtleNeckWidth,
	nEntitiesFound,
	nEntitiesRequired,
	nZipLength ;
BOOL
	bSnapCrotch,
	bBack,
   	bEOStoSelectedPoint,
  	bEOStoStandard,
  	bEOStoGivenLength,
	bFrontZipper,
	bBackZipper,
	bLoop ;
MSGID
	Ok;
	
ANGLE
	aAngle;
// Variables:
// ------------------------------------------ Start -------------------------------------------
XY	
	xyInt; 	// if it exist, intersection of line and line
NUMBER 
	nK1, nK2, nM1, nM2, nSlope1, nSlope2, nCase, 
	nY, nX, nLenghtToInt ;
ANGLE
	aAngleAtInt ;
	
// ------------------------------------------ End   -------------------------------------------

// Function:  	BOOLEAN = FN_LinLinInt( xyLineStart, xyLineEnd);
//
// Parameters:	xyLineStart1 = %1, %2
//		xyLineEnd1 = %3, %4
//		xyLineStart2 = %5, %6
//		xyLineEnd2 = %7, %8
//
// Returns:	%true if intersection found
//		%false if no intesection
//		xyInt =  intersection ( Global Variable )
//
		 

function FN_LinLinInt()
{
  // Caclculate slope
   nCase = 0 ;
   nSlope1 = Calc("angle", %1, %2, %3, %4);
   if ( nSlope1 ==  0 || nSlope1 == 180 ) nCase = nCase + 1 ;
   if ( nSlope1 ==90 || nSlope1 == 270 ) nCase = nCase + 2 ;

   nSlope2 = Calc("angle", %5, %6, %7, %8);
   if ( nSlope2 == 0  || nSlope2 == 180 ) nCase = nCase + 4 ;
   if ( nSlope2 ==90 || nSlope2 == 270 ) nCase = nCase + 8 ;

  // Parallel orthogonal lines, no intersection possible
   if ( nCase == 5 || nCase == 10) return (%false); 
   
  // Line1 is Horizontal and the Line2 is Vertical 
   if ( nCase == 9 ) {
   	nX = %5;
   	nY = %2;
	}
  // Line1 is Vertical and the Line2 is Horizontal 
   if ( nCase == 6 ) {
   	nX = %1;
   	nY = %6;
	}
	
  // Line 1 is Horizontal or Line 2 is horizontal 
     if (nCase == 1 || nCase == 4 ) {
      	if (nCase == 1) {
   		nM1 = (%8 - %6) / (%7 - %5) ;		// Slope
   		nK1 = %6 - nM1*%5;		// Y-Axis intercept
		nY = %2;
		}
	else{
   		nM1 = (%4 - %2) / (%3 - %1) ;		// Slope
   		nK1 = %2 - nM1*%1;		// Y-Axis intercept
		nY= %6;
		}
   		
   	// Solve for X at the given Y value
      	nX = (nY - nK1) / nM1;
	}
	
  // Line 1 is Vertical or Line 2 is Vertical 
	
     if (nCase == 2 || nCase == 8 ) {
      	if (nCase == 2) {
   		nM1 = (%8 - %6) / (%7 - %5) ;		// Slope
   		nK1 = %6 - nM1*%5;		// Y-Axis intercept
		nX = %1;
		}
	else{
   		nM1 = (%4 - %2) / (%3 - %1) ;		// Slope
   		nK1 = %2 - nM1*%1;		// Y-Axis intercept
		nX= %5;
		}
   		
   	// Solve for Y at the given X value
 	nY = (nM1*nX) + nK1;
	}

  // Both lines are Non-Orthogonal Lines
   if ( nCase==0  ) {
   	nM1 = (%4 - %2) / (%3 - %1) ;		// Slope
   	nM2 = (%8 - %6) / (%7 - %5) ;		// Slope
   	if (nM1 == nM2) return (%false) ;
   	nK1 = %2 - nM1*%1;		// Y-Axis intercept
   	nK2 = %6 - nM2*%5;		// Y-Axis intercept
   	if (nK1 == nK2) return (%false) ;
	// Find X
	nX = (nK2 - nK1) / (nM1 - nM2);	
	// Find Y				
 	nY = (nM1*nX) + nK1;
 	}

  // Ensure that the points X and Y are on the lines
  // Line 1 
   if ( nX < min(%1,%3) || nX > max (%1,%3)) return (%false) ;
   if ( nY < min(%2,%4) || nY > max(%2,%4)) return (%false) ;
   
  // Line 2
   if ( nX < min(%5,%7) || nX > max(%5,%7)) return (%false) ;
   if ( nY < min(%6,%8) || nY > max(%6,%8)) return (%false) ;
   
   xyInt.x = nX;
   xyInt.y = nY;
   
   return (%true);
   
   } // End function FN_LinLinInt


procedure PR_SetZipperID () {
	// Label entity with ID string and make Zipper %true
   	SetDBData(hEnt, "ID", sZipperID );
   	SetDBData(hEnt, "Zipper", "1" );
}


function FN_LineAndPolyInterSection () {  
   nLengthToInt = 0 ;      
   nVertex = GetVertexCount (%1) ;
   nn = 1; 
   GetVertex ( %1, nn, &xyTmp1) ; 
   while (nn <= nVertex) {
   	nn = nn + 1;
   	xyTmp = xyTmp1 ;
   	GetVertex ( %1, nn, &xyTmp1) ; 
   	if ( FN_LinLinInt (%2, %3, %4, %5, xyTmp, xyTmp1)) {
		nLengthToInt = nLengthToInt  + Calc ("length", xyTmp, xyInt) ;
		aAngleAtInt = Calc ("angle", xyTmp, xyInt) ;
   		return ( nLengthToInt ) ;	
		}
	else {
		nLengthToInt = nLengthToInt  + Calc ("length", xyTmp, xyTmp1) ;
		}
  	}
    return(nLengthToInt ) ;	
   } //End FN_LinePolyInterSection 
   
  function FN_LengthToPolyInterSection () {  
   nLengthToInt = 0 ;      
   nVertex = GetVertexCount (%1) ;
   nn = nVertex ; 
   GetVertex ( %1, nn, &xyTmp1) ; 
   while (nn >= 1) {
   	nn = nn - 1;
   	xyTmp = xyTmp1 ;
   	GetVertex ( %1, nn, &xyTmp1) ; 
	nLengthToInt = nLengthToInt  + Calc ("length", xyTmp, xyTmp1) ;
	if ( %2 <= nLengthToInt) {
		aAngleAtInt = Calc ("angle", xyTmp1, xyTmp) ;
		xyInt = CalcXY ("relpolar", xyTmp1, nLengthToInt - %2, aAngleAtInt) ;
		return(%true) ;	
		}	
	} 	
    return(%false) ;	
   } //End FN_LengthToPolyInterSection 
   
  	
// Find JOBST installed directory
// Set path to macros N.B. RE-USE OF CODE
//
   STRING	 sPathJOBST;
   sPathJOBST = GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;
   SetData("PathMacro",  sPathJOBST + "\\BODY" );
	
// Reset 
    UserSelection ("clear");
    UserSelection("update") ;
    Execute ( "menu", "SetStyle", Table( "find",  "style", "bylayer"));
    Execute ( "menu", "SetColor", Table( "find",  "color", "bylayer"));


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
			GetDBData(hEnt,"Data", s1);
			nSnapToFrontB = Value ("scalar", s1);
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
			GetDBData(hEnt,"Data", s2);
			nTurtleNeckWidth = Value ("scalar", s2);	
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
			nEntitiesRequired = 10 ;
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

   bEOStoSelectedPoint = %false ;
   bEOStoStandard = %false ;
   bEOStoGivenLength = %false ;
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

	if (bEOStoSelectedPoint || bEOStoStandard  )
		Loop = %false ;
	else {
		nZipLength = Value("scalar", sZipLength);
		if ( nZipLength == 0  && StringLength (sZipLength) > 0) {
			Display ("message", "error", "Invalid given length!\n\nTo use this option, type over the text in the \"EOS to:\" box with the required length in Inches (Decimal Inches).") ;
			Loop = %true ;
			}
   		else {
   			bEOStoGivenLength = %true ;
			Loop = %false ;
			}
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
  if (bEOStoStandard) {
	   if (bBackZipper){
		GetDBValue (hBack, "ZipperLength", &sTmp, &nZipLength );
		aTextAngle = 0;	// Always Horizontal
		xyText.x =xyBackW.x + (xyBackN.x- xyBackW.x)/2;
		xyText.y = xyBackW.y + .125;
		nLength = Calc("length", xyBackB, xyBackW) ;
		aAngle =  Calc("angle", xyBackB, xyBackW); 
		xyPt1 = CalcXY ("relpolar", xyBackB, nLength/2, aAngle );
		xyPt2 = xyPt1;
		xyPt2.y = xyPt1.y + 0.125;
		xyPt3 = CalcXY ("relpolar", xyPt1, 0.5, aAngle) ;	
		nZipLength = (nZipLength + nTurtleNeckWidth) - (nLength / 2);
		nZipLength = (nZipLength - 0.125) / 0.95 ;
   		}   	
	   else{
		GetDBValue (hFront, "ZipperLength", &sTmp, &nZipLength );
		//Temp use of nLength
		nLength = Calc("length", xyFrontW, xyFrontC) ;
		aTextAngle =  Calc("angle", xyFrontW, xyFrontC); 
		xyText = CalcXY ("relpolar", xyFrontW, nLength/2, aTextAngle ) ;
		xyText.y = xyText.y - .25;
		if (bSnapCrotch){
			nLength = 1.125 ;
			aAngle =  Calc("angle", xyFrontB, xyFrontW) ; 
			if (nSnapToFrontB  > 0) {
				xyPt1 = CalcXY ("relpolar", xyFrontB, (nLength - nSnapToFrontB) , aAngle ) ;		
				}
			else {
				xyPt1 = CalcXY ("relpolar", xySnap, nLength, aAngle ) ;
				}
			nZipLength= nZipLength -  Calc("length", xyFrontB, xyFrontW) ;
			nZipLength= nZipLength +  Calc("length", xyPt1, xyFrontW) ;
			} 
		else {
			nLength = Calc("length", xyFrontB, xyFrontW) ;
			aAngle =  Calc("angle", xyFrontB, xyFrontW); 
			xyPt1 = CalcXY ("relpolar", xyFrontB, nLength/2, aAngle ) ;
			nZipLength = nZipLength - (nLength / 2);
			}
		xyPt2 = xyPt1;
		xyPt2.y = xyPt1.y - 0.125;
		xyPt3 = CalcXY ("relpolar", xyPt1, 0.5, aAngle) ;	
		nZipLength = ((nZipLength+nTurtleNeckWidth) - 0.125) / 0.95 ;
   		}
   	}
 else {
	// A length has been given for the zipper
	//        
	if (bEOStoGivenLength ){
	   	bIntFound = %false ;
		nLength = (nZipLength - nTurtleNeckWidth) ;
	   	nLength = (nLength * 0.95) + 0.125;
		if (bFrontZipper ) {
			bIntFound = FN_LengthToPolyInterSection (hFront, nLength) ;
			xyPt2.y = xyInt.y - 0.125;
			xyPt1 = xyInt ;
			xyPt2.x = xyPt1.x ;
			xyPt3 = CalcXY ("relpolar", xyPt1, 0.5, aAngleAtInt) ;		   
			nLength = Calc("length", xyFrontW, xyFrontC) ;
			aTextAngle =  Calc("angle", xyFrontW, xyFrontC); 
			xyText = CalcXY ("relpolar", xyFrontW, nLength/2, aTextAngle ) ;
			xyText.y = xyText.y - .25;
			}
		else{	
			bIntFound = FN_LengthToPolyInterSection (hBack, nLength) ;
			xyPt2.y = xyInt.y + 0.125;
			xyPt1 = xyInt ;
			xyPt2.x = xyPt1.x ;
			xyPt3 = CalcXY ("relpolar", xyPt1, 0.5, aAngleAtInt) ;		   
			aTextAngle = 0;	// Always Horizontal
			xyText.x =xyBackW.x + (xyBackN.x- xyBackW.x)/2;
			xyText.y = xyBackW.y + .125;
			}
	        	if ( !bIntFound ) Exit (%cancel, "Can't find point on cutout line!\n") ;
	   }   

	// The zipper will be located at the X value of the given point
	//   
	if (bEOStoSelectedPoint ){
	   	if(!GetUser ("xy","Select zipper start point on Cut-Out Profile", &xyRequestedPoint)) Exit (%ok, "User cancelled") ;
		// get intersection at the given points X value on the front or back profile.	
		// Create construction points
		xyPt1 = CalcXY("relpolar", xyRequestedPoint, 100, 90) ;
		xyPt2 = CalcXY("relpolar", xyRequestedPoint, 100, 270) ;
	
		if (bFrontZipper ) {
			nZipLength= FN_LineAndPolyInterSection (hFront, xyPt1, xyPt2) ;
			xyPt2.y = xyInt.y - 0.125;
			nLength = Calc("length", xyFrontW, xyFrontC) ;
			aTextAngle =  Calc("angle", xyFrontW, xyFrontC); 
			xyText = CalcXY ("relpolar", xyFrontW, nLength/2, aTextAngle ) ;
			xyText.y = xyText.y - .25;
			}
		else{
			nZipLength= FN_LineAndPolyInterSection (hBack, xyPt1, xyPt2) ;
			xyPt2.y = xyInt.y + 0.125;
			aTextAngle = 0;	// Always Horizontal
			xyText.x =xyBackW.x + (xyBackN.x- xyBackW.x)/2;
			xyText.y = xyBackW.y + .125;
			}
        		if (nZipLength == 0 ) Exit (%cancel, "Can't find point on cutout line!\n") ;
		xyPt1 = xyInt ;
		aAngle = aAngleAtInt ;
		nZipLength = (nZipLength + nTurtleNeckWidth) ;
		nZipLength = (nZipLength - 0.125) / 0.95 ;
		xyPt2.x = xyPt1.x ;
		xyPt3 = CalcXY ("relpolar", xyPt1, 0.5, aAngle) ;	
	 }    	 
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
   hTextEnt = AddEntity ("symbol", "TextAsSymbol", xyText , 1, 1, aTextAngle) ;  

// Create ID string
   sZipperID =  sStyleID + MakeString("scalar", UID("get", hTextEnt )) ;
   SetDBData(hTextEnt, "ID", sZipperID );
   SetDBData(hTextEnt, "Zipper", "1" );
   sZipLength = Format("length", nZipLength);
   SetDBData(hTextEnt, "Data", sZipLength+" "+sZipType );
       	  
// Draw mark
   hEnt = AddEntity ("line", xyPt1, xyPt2) ;
   PR_SetZipperID() ;
   hEnt = AddEntity ("line", xyPt2, xyPt3) ;
   PR_SetZipperID() ;
	

// Reset and exit
   Execute ("menu", "SetLayer",  Table("find", "layer", "1")) ;
   Exit (%ok, "Zipper drawing Complete");

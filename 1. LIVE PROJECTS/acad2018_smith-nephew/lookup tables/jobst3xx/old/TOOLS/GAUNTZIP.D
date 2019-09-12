// File Name:	GAUNTZIP.D
// Purpose:		Draw zippers for ARM / Sleeves with gauntlets
// Version:		1.01
// Date:		19.Feb.98	
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
// 28 Sept 99	GG	Removed radiobuttons as they fail under NT

   Display ("cursor", "wait", "Gauntlet Zipper Tool");

HANDLE
	hDlg,
	hEnt,
	hThumbHoleArc,
	hGauntEnd,
	hLFS,	
	hChan,
	hCurve,
	hLayer;

XY
	xyZipStart,
	xyZipEnd,
	xyCen,
	xyThumbHoleArc,
	xyThumbHoleMax,
	xyMax, xyMin,
	xyTemplateMax, xyTemplateMin,
	xyGivenPoint,
	xyPt,
	xyPt1, 
	xyZipText,
	xyTmp, xyTmp1;

STRING
	sStyle,
	sZipText,
	sZipperType,
	sData,
	sAge,
	sElastic,
	sUnits,
	sZipLength,
	sZipOffset,
	sProfileID,
	sTmp, sTmp1,
	sSide,
	sDlgElasticList,
	sDlgLengthList,
	sDlgZipperList,
   	sGauntEnd,
   	sOriginMark,
  	sThumbHoleArc,
	sZipperID; // Unique ID for each Zipper, allows easy selection and deletion

NUMBER
	Ok,
	nFound,
	nRadius,
	nThumbHoleRadius,
	nSeam,
	nAge,
	nType,
	nReqired,
	nStandardCutBack,
	nProfileLength,
	nFoldLength,
	nElasticFactor,
	nElastic,
	ii, nn,
	nFoldZipperMax ,
	nProfileZipperMax ,
	nZipLength ;
	
BOOL
	EOStoSelectedPoint,
	EOStoGivenLength,
	EOStoStandardPoint,
	bxyThumbFound,
	Loop,
	bDetachable,
	bLateral,
	bPalmer,
	bMedial,
	bDorsal,
	bLFS;

ANGLE
	aStart,
	aZipper,
	aStartMarker,
	aEndMarker,
	aDelta,
	aAngle;

// Clear user selection
//
   UserSelection ("clear") ;
   UserSelection ("update") ;
   
   STRING		sPathJOBST;
   sPathJOBST = GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;
   SetData("PathMacro",  sPathJOBST + "\\TOOLS" );
   

// ************************   
// Prompt user selection
// ************************   
   if ( !hCurve=GetEntity ("Select Arm or Sleeve Profile", "all")) Exit (%ok, "User Cancelled") ;

   Display ("cursor", "wait", "Searching for more data");
      
   GetDBData (hCurve, "ID", &sProfileID);
   nStringLength = StringLength(sProfileID);
   if (StringCompare( "LeftProfile", StringMiddle(sProfileID, nStringLength - 10, 11))){
	sSide = "Left";
	sStyle = StringMiddle(sProfileID, 1, nStringLength - 11) + sSide;
	}
   if (StringCompare( "RightProfile", StringMiddle(sProfileID, nStringLength - 11, 12))) {
	sSide = "Right";
	sStyle = StringMiddle(sProfileID, 1, nStringLength - 12) + sSide;
	}


// Check if this is a valid sleeve or arm Gauntlet, exit if not!
// This is crude check only
    if (StringLength (sSide) == 0) Exit (%cancel, "A Sleeve or Arm profile was not selected\n") ;
    if (!StringCompare(StringMiddle(sProfileID, 1,2), "GT")) Exit (%cancel, "The selected profile is not for a GAUNTLET!\n") ;

   		   		
// Ensure that the arm profile is horizontal
// This provides for much simpler code than catering for a none horizontal case
   GetExtent (hCurve, &xyMin, &xyMax) ;
   if (abs(xyMin.y - xyMax.y) > abs(xyMin.x - xyMax.x)) Exit (%cancel, "The selected profile is not HORIZONTAL.\nThis zipper tool only works if the Sleeve / Arm is rotated to a horizontal orientation.") ;

// ************************   
// Get data from drawing
// ************************   
   sGauntEnd = sStyle + "GauntEnd" ;
   sOriginMark = sStyle + "originmark" ;
   sThumbHoleArc = sStyle + "ThumbHole" ;
   
   sTmp = "DB ID = '" + sGauntEnd  +  "' OR DB ID = '" + sOriginMark + "' OR (DB ID = '" + sThumbHoleArc + "' AND type='arc')"   ;
   xyThumbHoleArc.y = -10000;		// Dummy value
   xyThumbHoleMax.x = -10000; 	//	"	"
   	     
   nFound = 0 ; nReqired = 5;
   hChan = Open ("selection", sTmp );
   if (hChan) {
	ResetSelection(hChan);
	while (hEnt = GetNextSelection(hChan)) {
		sTmp = "" ;
		GetDBData(hEnt,"ID", &sTmp) ;
		if (StringCompare(sTmp, sGauntEnd )){
			// End Of Gauntlet
			nFound = nFound + 1 ;
			hGauntEnd = hEnt ;
			}
		if (StringCompare(sTmp, sOriginMark )){
			// Origin marker
			nFound = nFound + 1 ;
			GetDBData(hEnt, "Age", &sAge);
			nAge = Value("scalar", sAge) ;
			}
		if (StringCompare(sTmp, sThumbHoleArc )){
			// Locate lower arc of Thumb hole Gauntlet
			nFound = nFound + 1 ;	// Note we find all 3 even if we only use the bottom left arc				
			GetGeometry(hEnt, &xyCen, &nRadius, &aStart, &aDelta) ;
			GetExtent( hEnt, &xyMin, &xyMax);
			if (xyMax.x > xyThumbHoleMax.x ) {
				// This gets the arc at the extreme X (right)
				xyThumbHoleMax.x = xyMax.x;
				xyThumbHoleMax.y = xyMin.y + (xyMax.y - xyMin.y)/2;
				}
			if (xyCen.y > xyThumbHoleArc.y) {
				// The arc we are looking for always has the highest centre
				hThumbHoleArc = hEnt ;
				xyThumbHoleArc = xyCen ;
				nThumbHoleRadius = nRadius ;
				}
			}
		}
   	}

   Close("selection", hChan); 

// Check that sufficent data have been found, otherwise exit
//
    if (nFound < nReqired) Exit (%cancel, "Missing data for selected Glove!\n") ;
    if (nFound > nReqired) Exit (%cancel, "Two or more drawings of the same style exist!\nDelete the extra drawing/s and try again.\n") ;
    
// Load Standard Functions
//
   @ZIPDEF.D;
   
// Define local functions 
// Name:	FN_PolyInterSection ( hPoly, nLength, nXvalue)
// Type:	Function 
// Returns:	%false 	if no intersection found.
//		nLength	if an intersection is found then it retuns the length to
//				the intesection.
//		xyInt		The intesection point.
//
// Args:	hPoly		%1, Handle of the polyline curve (Vertex MUST be in order of 
//				increasing X ie  X1 < Xn.
//		nLength	%2, If nLength > 0 then find the intesection at the given length.
//		nXvalue  	%3, If nXvalue > 0 then intesect at the given value for X.
//
//NB:		if nLength > 0 then nXvalue = 0; if nXvalue > 0 then nLength = 0.
//		nLength <> nXvalue.
//		Function is only valid in +ve Y, +ve X, quadrant
// 		
function FN_PolyInterSection () {        
   nVertex = GetVertexCount (%1) ;
   nTraversed = 0 ;
   nn = 1; 
   GetVertex ( %1, nn, &xyTmp1) ; 
   if (%3 > 0) {
   	GetExtent ( %1, &xyPt, &xyPt1) ;
   	xyPt.x = %3 ;
   	xyPt1.x = %3 ;
    	while (nn <= nVertex) {
   		nn = nn + 1;
   		xyTmp = xyTmp1 ;
   		GetVertex ( %1, nn, &xyTmp1) ; 
   		if ( FN_LinLinInt (xyPt,xyPt1, xyTmp, xyTmp1)) {
			nTraversed = nTraversed + Calc ("length", xyTmp, xyInt) ;
    			return ( nTraversed) ;	
			}
		nTraversed = nTraversed + Calc ("length", xyTmp, xyTmp1) ;
  		}
   	}
   if (%2 > 0) {
    	while (nn <= nVertex) {
   		nn = nn + 1;
   		xyTmp = xyTmp1 ;
   		GetVertex ( %1, nn, &xyTmp1) ; 
   		nLength = Calc ("length", xyTmp, xyTmp1) ;
		if ((nTraversed + nLength) >= %2) {
			nLength = %2 - nTraversed ;
			if (!FN_CirLinInt( xyTmp, xyTmp1, xyTmp, nLength)) return ( %false) ;	
    			return ( %2) ;	
			}
		else
			nTraversed = nTraversed + nLength ;
  		}
   	}     
   return(%false) ;	
   
   } //End FN_PolyInterSection 
        
// Setup diff. between Sleeve/Arm Gauntlet and the detachable gauntlet
//   
   if (StringCompare(StringMiddle(sProfileID, 1,4), "GTGT")) {
   	sDlgElasticList =  "3/8\" Elastic\n3/4\" Elastic\n1½\" Elastic\nNo Elastic";
   	bDetachable = %true ;
   	if (nAge < 10)
   		sDlgElasticList =  "3/8\" Elastic\n" + sDlgElasticList;  
   	else
      		sDlgElasticList =  "3/4\" Elastic\n" + sDlgElasticList; 
   	}
   else {
         	sDlgElasticList =  "No Elastic" ;
    	bDetachable = %false ;   
        	}
   
   nStandardCutBack = 1.375 ;
   
   sDlgZipperList = "Lateral\nPalmer\nMedial\nDorsal\nL.F.S" ;
   
// Calculate the extents of the template and the maximum zipper length.
// These will then be used to verify the input data and subsequently in drawing
// the zippers on the template.
//
   GetExtent (hCurve, &xyTmp1, &xyTemplateMax) ;
   GetExtent (hGauntEnd, &xyTemplateMin, &xyTmp) ;
   nGauntletExtension =  xyTmp1.x - xyTemplateMin.x ;
   
   //Maximum zipper length on the fold
   nFoldLength = (xyTemplateMax.x - xyTemplateMin.x)  ;
   nFoldZipperMax = nFoldLength  - nStandardCutBack ;

   //Maximum zipper length along the profile
   SetDBData(hCurve, "ZipperLength", "%length") ;
   GetDBValue(hCurve, "ZipperLength", &sTmp1, &nProfileLength );
   nProfileZipperMax = (nProfileLength + nGauntletExtension) - nStandardCutBack ;      
   
// **********************   
// Get data from user
// **********************   

// Create Dialog
// Length specification      	
   sDlgLengthList =  "Standard\nGive a length\nSelected Point";

   Loop = %true ;
   while (Loop) {		
   	EOStoSelectedPoint = %false ;
   	EOStoStandardPoint = %false ;
   	EOStoGivenLength = %false ;  
   
          nButX = 65; nButY = 60;

                    hDlg = Open ("dialog", sSide + " Gauntlet Zipper", "font Helv 8", 20,20, 210, 80);
 
	AddControl ( hDlg, "pushbutton",  nButX, nButY, 35, 14, "Cancel", "%cancel", "");
	AddControl ( hDlg, "pushbutton",  nButX + 48 , nButY, 35,14, "OK", "%ok", "");
	
//	AddControl(hDlg, "radiobutton", 15, 2, 60, 14, "Lateral Zipper", "number", "bLateral") ;
 	AddControl ( hDlg, "dropdownlist", 15, 12, 70, 70 , sDlgZipperList, "string", "sZipperType");
//	AddControl(hDlg, "radiobutton", 15, 12, 60, 14, "Palmer Zipper", "number", "bPalmer") ;
//	AddControl(hDlg, "radiobutton", 15, 22, 60, 14, "Medial Zipper", "number", "bMedial") ;
//	AddControl(hDlg, "radiobutton", 15, 32, 60, 14, "Dorsal Zipper", "number", "bDorsal") ;
//	AddControl(hDlg, "radiobutton", 15, 42, 60, 14, "L.F.S Zipper", "number", "bLFS") ;
	
	AddControl ( hDlg, "ltext", 106, 12, 10, 14 , "To:", "string", "");
	AddControl ( hDlg, "combobox", 125, 10, 70, 40 , sDlgLengthList , "string", "sZipLength");

	AddControl ( hDlg, "ltext", 95, 34, 30, 14 , "Elastic:", "string", "");
 	AddControl ( hDlg, "dropdownlist", 125, 32, 70, 70 , sDlgElasticList, "string", "sElastic");

      	Ok = Display("dialog", hDlg, "%center");
      	
  	Close("dialog", hDlg);

	if (Ok == %cancel ) Exit (%ok, "User Cancel!") ;	
	
	bLateral = %false; bPalmer = %false; bMedial = %false; bDorsal = %false;  bLFS = %false;
	
	if (StringCompare(sZipperType, "Lateral" )) bLateral = %true ;
	if (StringCompare(sZipperType, "Palmer" )) bPalmer = %true ;
	if (StringCompare(sZipperType, "Medial" )) bMedial = %true ;
	if (StringCompare(sZipperType, "Dorsal" )) bDorsal = %true ;
	if (StringCompare(sZipperType, "L.F.S" )) bLFS = %true ;
	
//	if (!bLateral && !bPalmer && !bMedial && !bDorsal && !bLFS) {
//		Display ("message", "error", "Select a Zipper type!") ;	
//		continue ;
//		}
		
	if (StringCompare("Selected Point", sZipLength)) EOStoSelectedPoint = %true;
	
	if (StringCompare("Standard", sZipLength)) EOStoStandardPoint = %true;
	    
	if ( EOStoSelectedPoint || EOStoStandardPoint )
		Loop = %false ;			
	else {
		nZipLength = Value("scalar", sZipLength);
		if ( nZipLength == 0  && StringLength (sZipLength) > 0) {
			Display ("message", "error", "Invalid given length!\n\nTo use this option, type over the text in the \"To:\" box with the required length in Inches (Decimal Inches).") ;
			Loop = %true ;
			}
   		else {
   			if ( ((bPalmer || bMedial) && nZipLength > nProfileZipperMax) || ((bDorsal || bLFS || bLateral) && nZipLength > nFoldZipperMax)) {
				Display ("message", "error", "Requested length too long!\nThe zipper as requested will not end on the template. Use a shorter zipper.") ;
				Loop = %true ;
   				} 
   			else {
				Loop = %false ;
				EOStoGivenLength = %true ;  
				}	
			}
		}
	if ((EOStoGivenLength || EOStoSelectedPoint ) && bDetachable && !StringCompare(StringMiddle(sElastic,1,1), "N")) {
		Display ("message", "error", "There should be no elastic for a detachable gauntlet if the Zipper does not end at the EOS!\nTry again, selecting \"no elastic\"") ;	
		Loop = %true ;
		continue ;	
		}
        }   
    
   Display ("cursor", "wait", "Drawing Zipper");
	    	
// Establish allowance for zippers
//
   nElastic = 0.75 ;
   if (StringCompare(StringMiddle(sElastic,1,1), "N")) nElastic = 0 ;
   if (StringCompare(StringMiddle(sElastic,1,3), "3/8")) nElastic = 0.375 ;
   if (StringCompare(StringMiddle(sElastic,1,1), "1")) nElastic = 1.5 ;
   
   nElasticFactor = 0.92 ;
   
// **********************   
// Zipper START POINT
// **********************   
   GetExtent( hGauntEnd, &xyMin, &xyMax);
   
   nSeam = 0.1875 ;	
   if (bDetachable) {
   	// X values
   	xyZipStart.x = xyMin.x + nStandardCutBack   ;
      	// Y values
   	if (bDorsal || bLateral) xyZipStart.y = xyMin.y + nSeam ;
   	if (bPalmer || bMedial) {
  		if ( 0 == FN_PolyInterSection ( hCurve, 0, xyZipStart.x)) 
  			xyZipStart.y = xyMax.y ;  
  		else
  			xyZipStart = xyInt;   		
   		}
   	if (bLFS){
   		GetExtent( hThumbHoleArc, &xyMin, &xyMax);
   		xyZipStart.y = xyMax.y ;
		if ( xyZipStart.x > xyMax.x) xyZipStart.y = xyMin.y ;
      		if (FN_CirLinInt (xyZipStart, xyZipStart.x, xyZipStart.y - 5, xyThumbHoleArc, nThumbHoleRadius)) xyZipStart = xyInt ;
   		}
   	}
   else { 
   	// X values
   	xyZipStart.x = xyMin.x ;
      	// Y values
   	if (bDorsal || bLateral) xyZipStart.y = xyMin.y + nSeam ;
   	if (bPalmer || bMedial) xyZipStart.y = xyMax.y ;  	
   	if (bLFS) {
   		GetExtent( hThumbHoleArc, &xyMin, &xyMax);
   		xyZipStart.y = xyMax.y ;
   		}
	}   

// *******************   
// Zipper END Point
// *******************         
// If a required point then get the point.
// Check that it will lie on the template (we can't be too accurate here)
// 
   
   if (EOStoSelectedPoint)  {
        	Loop = %true;
   	while(Loop ){
  		if (!GetUser("xy", "Select End of Zipper", &xyGivenPoint)) Exit(%ok, "Start not selected") ;
		if ( xyGivenPoint.x < xyTemplateMin.x ||  xyGivenPoint.x > xyTemplateMax.x ) {
			Ok = Display("message", "OKquestion", "Given point can't be used! Try again.\nOr use cancel to Exit");
			if ( Ok == %cancel) Exit (%ok, "User Cancelled");
			}
		else
			Loop = %false;
		}	
   	Display ("cursor", "wait", "Drawing Zipper");
	}
   
   if (bPalmer || bMedial ) {
//      	if (bDetachable)nGauntletExtension = 0 ;
   	aStartMarker = 315;
   	aEndMarker = 225 ;
  	if (EOStoSelectedPoint){
  		if (!nDrawnLength = FN_PolyInterSection ( hCurve, 0, xyGivenPoint.x)) Exit (%abort, "Unable to find intersection on profile!");;
    		nZipLength = (nDrawnLength + nGauntletExtension + nElastic ) / nElasticFactor ;
  	 	}
  	if (EOStoStandardPoint){
  		nDrawnLength = nProfileLength - nStandardCutBack ;
  		if (!bDetachable){
  			if (!nDrawnLength = FN_PolyInterSection ( hCurve, nDrawnLength , 0))  Exit (%abort, "Unable to find intersection on profile!");;
    			}
    		else	{
   			nVertex = GetVertexCount (hCurve) ;
   			GetVertex (hCurve, nVertex , &xyInt ) ; 
    			}
    		nZipLength = (nDrawnLength + nGauntletExtension + nElastic ) / nElasticFactor ;
  	 	}
  	if (EOStoGivenLength){
      		nDrawnLength =((nZipLength * nElasticFactor ) - nElastic) - nGauntletExtension  ; 
  		if (!nDrawnLength = FN_PolyInterSection ( hCurve, nDrawnLength , 0))  Exit (%abort, "Unable to find intersection on profile!");;  		 	
  	 	}
  	xyZipEnd = xyInt ;		// From FN_PolyInterSection above
   	}             

   if ( bLateral || bDorsal || bLFS ) {
      	if (!bDetachable) nGauntletExtension = 0;	//N.B. NOT Detachable
      	if (bLFS) {
   		if (!bDetachable)
   	   		aStartMarker = 0;
   	   	else
   	   		aStartMarker = 225;   	   		
   		aEndMarker = 180 ;
   		xyZipEnd.y = xyThumbHoleMax.y ;
   		}
   	else {
   		aStartMarker = 45;
   		aEndMarker = 135 ;
   		xyZipEnd.y = xyZipStart.y ;
   		}
  	if (EOStoSelectedPoint){
  		nDrawnLength = xyGivenPoint.x - xyZipStart.x ;
    		nZipLength = (nDrawnLength + nElastic ) / nElasticFactor ;
  	 	}
  	if (EOStoStandardPoint){
  		nDrawnLength =  nFoldLength - nStandardCutBack  ;
    		nZipLength = (nDrawnLength + nElastic ) / nElasticFactor ;
   	 	}
  	if (EOStoGivenLength){
      		nDrawnLength =((nZipLength * nElasticFactor ) - nElastic)  - nGauntletExtension; 
  	 	}
	xyZipEnd.x =  xyZipStart.x + nDrawnLength  ;  		
   	}             
   	
   
//************** 
// Draw Zipper
//**************
// Each entity is given an ID Data Base value that links them together
// This is based on the UID of the symbol "ZipperText" and the Glove Profile ID.
// A Symbol is used as DB values can't be attached to text entities 

// Draw on layer Notes
   Execute ("menu", "SetLayer",  Table("find", "layer", "Notes")) ;

   sZipText = Format("length", nZipLength);
   aZipper = 0;
   xyZipText = CalcXY ("average", xyZipStart, xyZipEnd) ;
   xyZipText.y = xyZipText.y + 1.0 ;
   
   if (bLateral) sZipText = sZipText + " LATERAL ZIPPER" ;
   if (bLFS) {
   	sZipText = sZipText + " L.F.S. ZIPPER" ;
   	xyZipText.y = xyZipText.y - .75 ;   	
    	}
   if (bDorsal) sZipText = sZipText + " DORSAL ZIPPER" ;
   if (bPalmer){
    	sZipText = sZipText + " PALMER ZIPPER" ;
    	xyZipText.y = xyZipText.y - 2.5 ;
    	}   	
   if (bMedial){
    	sZipText = sZipText + " MEDIAL ZIPPER" ;
    	xyZipText.y = xyZipText.y - 2.5 ;
    	}   	
   
   SetSymbolLibrary( sPathJOBST + "\\JOBST.SLB") ;   	     
   Symbol( "find", "TextAsSymbol") ;
   hEnt = AddEntity ("symbol", "TextAsSymbol", xyZipText, 1, 1, aZipper) ;  
   SetDBData(hEnt, "Data", sZipText );
   
// Create ID string
   sZipperID =  sProfileID + MakeString("scalar", UID("get", hEnt)) ;
   
// Label entity with ID string and make Zipper %true
   SetDBData(hEnt, "ID", sZipperID );
   SetDBData(hEnt, "Zipper", "1" );
    	
// Add label and arrows
   hEnt = AddEntity("marker","closed arrow", xyZipStart , 0.5, 0.125, aStartMarker) ;	
   SetDBData(hEnt, "ID", sZipperID );
   SetDBData(hEnt, "Zipper", "1" );
   
   hEnt = AddEntity("marker","closed arrow", xyZipEnd , 0.5, 0.125, aEndMarker) ;		
   SetDBData(hEnt, "ID", sZipperID );
   SetDBData(hEnt, "Zipper", "1" );
   if (bLFS) {
   	hEnt = AddEntity("line",xyThumbHoleMax, xyZipEnd) ;		
   	SetDBData(hEnt, "ID", sZipperID );
   	SetDBData(hEnt, "Zipper", "1" );
   	}
   
// Reset and exit
   Execute ("menu", "SetLayer",  Table("find", "layer", "1")) ;
   Exit (%ok, "Zipper drawing Complete"); 
// FileName:	BODYDEF.D
// Purpose:	To Load the requisite functions and procedures
//		
// Version:	1.02
// Author:		Gary George
// Date:		3.Nov.93
// Called By:	BODY.D
//

//
// Function:  	number = FN_CirLinInt( xyLineStart, xyLineEnd, xyCircleCen, nRadius);
//
// Parameters:	xyLineStart = %1, %2
//		xyLineEnd = %3, %4
//		xyCirleCen = %5, %6
//		Radius = %7
//
// Returns:	%true if intersection found
//		%false if no intesection
//		xyInt = First intersection ( Global Variable )
//
		 
// Variables:
// ------------------------------------------ Start -------------------------------------------
XY	
	xyInt; 	// if it exist, 1st intersection of line and circle
NUMBER 
	nK, nA, nB, nC, nM, nCalcTmp, nRoot, nSlope, nSign;
// ------------------------------------------ End   -------------------------------------------

function FN_CirLinInt()
{
//Note - Returns only the First intersection

   nSlope = Calc("angle", %1, %2, %3, %4);

// Horizontal Line
   if ( nSlope == 0 || nSlope == 180 ) {
	nSlope = -1;
	nC = %7^2 - (%2-%6)^2;
   	if ( nC < 0 ) return(%false);		// No Roots
	nSign = 1;			// Test each root
                while ( nSign > -2){
		nRoot = %5 + sqrt(nC) * nSign;		
  		if ( nRoot >= min(%1,%3) && nRoot <= max(%1,%3))  {
			xyInt.x = nRoot;
			xyInt.y = %2;
			return (%true);
			}
                	nSign = nSign - 2;
		}
	return(%false);
	}

// Vertical line
   if ( nSlope==90 || nSlope == 270 ) {
	nSlope = -1;	
	nC = %7^2 - (%1-%5)^2;
	if ( nC <0 ) return(%false);		// No Roots
	nSign = 1;			// Test each root
	while ( nSign > -2){
		nRoot = %6 + sqrt(nC) * nSign;
  		if ( nRoot >= min(%2,%4) && nRoot <= max(%2,%4)){
			xyInt.y = nRoot;
			xyInt.x = %1;
			return (%true);
			}
		nSign = nSign -2;
		}
	return(%false);
	}

// Non-Orthogonal Line
   if ( nSlope > 0 ) {
   	nM = (%4 - %2) / (%3 - %1) ;		// Slope
   	nK = %2 - nM*%1;			// Y-Axis intercept
   	nA = (1 + nM^2) ;			
   	nB = 2 * ( -%5 + nM*nK - %6*nM) ;
  	nC = %5^2 + nK^2 + %6^2 - 2*%6*nK - %7^2 ;
   	nCalcTmp = (nB^2) - (4 * nC *nA);

   	if ( nCalcTmp < 0 ) return (%false);	// No Roots
	nSign = 1;
	while ( nSign > -2) {
   		nRoot = ( -nB + (sqrt(nCalcTmp)/nSign))/(2*nA);
  		if ( nRoot >= min(%1,%3) && nRoot <= max(%1,%3)) {  
			xyInt.x = nRoot;
	   		xyInt.y = nM * nRoot + nK;
			return (%true);
			}
		nSign = nSign -2;
		}
	return(%false);
 	}

   } // End function FN_CirLinInt
   
   
   
   
// Calculate the required vestcurve angle.
// Open the VESTCURVE data file
// Check for 1. Existance and 2. Corruption  
//
// Function	aAngle = FN_CurveAngle ( xyStart, nRadiusFromStart)
//
// Purpose	Get the angle of the curve w.r.t. the rotation of
//		of the vest curve
//
function FN_CurveAngle () {
   hCurve = Open ("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly");
   if (hCurve) { 
	if (!GetLine(hCurve, &sLine)) {
	   	Close ("file", hCurve);
		Exit (%abort, "Can't read VESTCURV.DAT\nFile maybe corrupted");
		}
	if (!ScanLine( sLine, "blank", &nLength, &aAngle)) {
   		Close ("file", hCurve);
		Exit (%abort, "Can't read VESTCURV.DAT\nFile maybe corrupted");
		}
	}
   else
	Exit (%abort, "Can't open VESTCURV.DAT\nCheck installation");
	
// Intersect the unrotated vest curve with the given radius, that is centered on the 
// xyAxilla Point, to get curve angle.
// This angle will be used to calculate vestcurve rotation.
//
// 
   aCurveRotation = 100000 ;	// Impossible value used to test for non-intersecton
   aPrevAngle = 0;
   xyPt1.x = %1;
   xyPt1.y = %2;

   while (GetLine(hCurve,&sLine)) {
	aAngle = aAngle + aPrevAngle;
	xyPt2 = CalcXY("relpolar", xyPt1, nLength, aAngle);
	if ( FN_CirLinInt( xyPt1, xyPt2, %1, %2, %3)){
		aCurveRotation = Calc ("angle", %1, %2, xyInt);  // V.I.P.
		break;
		}
	xyPt1 = xyPt2;
	aPrevAngle = aAngle;
	ScanLine(sLine, "blank", &nLength, &aAngle);	
  	}
  	
   Close ("file", hCurve);
   
   if (aCurveRotation == 100000) 
	Display("message","error","Can't Find Curve angle with this data in \"FN_CurveAngle\" ");
	
   return (aCurveRotation ) ;	   

} // End function
      

// Function:
//	nRounded Value = FNRound ( nLength )
// Returns:
//	nRounded Value	- The result rounding. 
//			  Based on the "measurement" options set in DRAFIX

// Arguments:
//	nLength

function FNRound () {
	return ( Value("length", Format("length", %1)));
} /* End FNRound */


// Layer setting procedures
//

procedure PRSetLayerTemplateRight () {
   hLayer = Table("find", "layer", "TemplateRight") ;
   if ( hLayer > %zero ) Execute ("menu", "SetLayer",  hLayer) ;
}

procedure PRSetLayerTemplateLeft () {
   hLayer = Table("find", "layer", "TemplateLeft") ;
   if ( hLayer > %zero ) Execute ("menu", "SetLayer",  hLayer) ;
}


procedure PRSetLayerConstruct () {
   hLayer = Table("find", "layer", "Construct") ;
   if ( hLayer > %zero ) Execute ("menu", "SetLayer",  hLayer) ;
}

procedure PRSetLayerNotes () {
   hLayer = Table("find", "layer", "Notes") ;
   if ( hLayer > %zero ) Execute ("menu", "SetLayer",  hLayer) ;
}



// Procedure PRDataStamp ( Origin.x, Origin.y, nStampType)
// Purpose:	To add a stamp containing the patient details 
//		at the given origin
//
// Arguments:	%1 = Origin.x
//		%2 = Origin.y
//		%3 = nStampType where :-
//			1    = Body, full stamp
//			11 = Body, Axilla insert left abbreviated stamp
//			12 = Body, Axilla insert right abbreviated stamp
//			2    = Sleeve, left full stamp
//			3    = Sleeve, right full stamp
//			4    = Work Order stamp
//			
// Notes:		The origin is the top of the text
//		text is printed downwards
	
	
procedure PRDataStamp () {
   SetData("TextVertJust", 32);		// Bottom
   SetData("TextHorzJust", 1);		// Left
   SetData("TextHeight", 0.125);
   SetData("TextFont", 0);		//CAD Block
   SetData("TextAspect" , 0.6) ;
   
//Store original layer and set to layer Notes
   GetData( "LayerNumber", &hOriginalLayer) ;
   PRSetLayerNotes () ;

   if ( %3 == 1 ) {
  	AddEntity("text", sPatient + "\n" + sWorkOrder + "\n" + sFabric  , %1,%2);
  	%2 = %2 - 0.8; 
   	PRSetLayerConstruct() ;
  	AddEntity("text", sFileNo + "\n" + sDiagnosis + "\n" + sAge  + "\n" + sSEX, %1,%2);
    } // End if 1 or  2 or 3 ( body, Left and right sleeves )

// Axilla Inserts stamp
//
    if ( %3 == 11 || %3 == 12) {
   	if ( %3 == 11)   AddEntity("text", "Vest Left" + "\n" +  sPatient + "\n" + sWorkOrder ,%1 ,%2);
		else  AddEntity("text", "Vest Right" + "\n" +  sPatient + "\n" + sWorkOrder ,%1, %2);
    }
    
   if ( %3 == 10)   AddEntity("text", "Vest" + "\n" +  sPatient + "\n" + sWorkOrder ,%1 ,%2);
   
// Restore original layer     
   Execute ("menu", "SetLayer",  hOriginalLayer) ;

} // End PRDataStamp() 


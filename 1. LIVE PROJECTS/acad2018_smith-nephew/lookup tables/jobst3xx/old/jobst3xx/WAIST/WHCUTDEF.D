// File:		WHCUTDEF.D
// Purpose:	To load Waist Height
//		Drawing functions
//		And string functions
// Version:	1.01
// Date:		9.Jan.94
// Author:	Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action

   sSpecificationNO	= "GOP 01-02/12" ;
   sSpecificationDate = "15 October 91" ;

   nSeam	= 0.1875;
   nFilletRadius =   0.1875;

   nStandardRedFac	= 0.83 ; 
   nBabyRedFac	= 0.88 ;
   nChildRedFac 	= 0.86 ;
   nUpliftRedFac 	= 0.88 ;
   nReductionTol	= 0.05 ;

// Template extension positioning on leg
// Values for each case
   nExtTempltTol1max	= 12.0;		// as mesured
   nExtTemplt1tape		=  -1;	

   nExtTempltTol2min	= 12.0 ;		// at 28 1/2" scale	
   nExtTempltTol2max	= 23.875 ;
   nExtTemplt2tape		= 25;	

   nExtTempltTol3min	= 23.875 ;	// at  27" scale	
   nExtTemplt3tape		= 24;	
 
// Crotch cutout factors, for front only, back = 1 - FrontFactor
//
   nFemaleOpenCrotchFac 	= 0.5 ;
   nFemaleClosedCrotchFac 	= 0.5 ;
   nChildTo3OpenCrotchFac 	= 0.5 ;
   nChildTo3ClosedCrotchFac 	= 0.5 ;
   nMaleOpenCrotchFac 	= 0.5 ;
   nMaleClosedCrotchFac 	= 0.4;
   nBodyFrontStep	 	= 1.0 ; 
   nBodyFrontIncreaseOff	= 0.25 ;
   nBodyFrontReduceOff	= 0.125;
   nCutOutDiaMaxTol	= 6.0 ;
   nCutOutConstructOff_1	= 0.75;
   nCutOutConstructFac_1	= 0.87 ;
   nBodyBackCutOutMinTol	= 3.0 ;
   nEndBackBodyOff	= 1.25 ;
   nPantyThighRed1		= 0.85;
   nPantyThighRed2 	= 0.92;

// Set up units conversion factor
   if ( StringCompare(sUnits, "cm")) nUnitsFac = 10/25.4;	// Cm to Inches   
   if ( StringCompare(sUnits, "inches")) nUnitsFac = 1;		// Inches


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




// Function:
//	nDecimalised = FNDecimalise( nValue)
// Returns:
//	nDecimalised	- The result of converting from the display format in the
//			  dialogue box based on the units type
//			  Eg. 
//			  If data is stored as 10.1 in "inches" then this means
//			  10" and 1/8"   ie inches and eighths
// Arguments:
//	nValue
// Global Variables
//	sUnits 		- The input units "cm" | "inches"
//	nInt		- The value befor the decimal point
//	nDec		- The value after the decimal point
//	sError		- Error Reporting string
//
function FNDecimalise() {
if (StringCompare("inches", sUnits)) {
	nInt = int(%1);
	nDec = ( %1 - nInt) ; 
	nDec = nDec * 1.25 ;
	return (nInt + nDec) ;
	}
else	
	return ( %1) ;
} /* End FNDecimalise */

// Layer setting procedures
//



procedure PRSetLayerTemplateRight () {
   hLayer = Table("find", "layer", "TemplateRight") ;
   if ( hLayer > %zero )
 	Execute ("menu", "SetLayer",  hLayer) ;
   else
   	Table ("add", "layer", sLayer, "Paper Template Right Hand side", "current", Table("find", "color", "Red"));
}

procedure PRSetLayerTemplateLeft () {
   hLayer = Table("find", "layer", "TemplateLeft") ;
   if ( hLayer > %zero )
 	Execute ("menu", "SetLayer",  hLayer) ;
   else
   	Table ("add", "layer", sLayer, "Paper Template Left Hand side", "current", Table("find", "color", "Blue"));
}


procedure PRSetLayerConstruct () {
   hLayer = Table("find", "layer", "Construct") ;
   if ( hLayer > %zero )
 	Execute ("menu", "SetLayer",  hLayer) ;
   else
   	Table ("add", "layer", sLayer, "Construction layer", "current", Table("find", "color", "Gray"));
}

procedure PRSetLayerData () {
   hLayer = Table("find", "layer", "Data") ;
   if ( hLayer > %zero )
 	Execute ("menu", "SetLayer",  hLayer) ;
   else
   	Table ("add", "layer", sLayer, "Data Layer", "current", Table("find", "color", "Black"));
}

procedure PRSetLayerNotes () {
   hLayer = Table("find", "layer", "Notes") ;
   if ( hLayer > %zero )
 	Execute ("menu", "SetLayer",  hLayer) ;
   else
   	Table ("add", "layer", sLayer, "Notes in Green", "current", Table("find", "color", "Green"));
}




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

// File:		WHFTRDEF.D
// Purpose:	To load Leg Reinforcement
//		Drawing functions
//		And string functions
// Version:	1.01
// Date:		07.Feb.94
// Author:	Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action

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


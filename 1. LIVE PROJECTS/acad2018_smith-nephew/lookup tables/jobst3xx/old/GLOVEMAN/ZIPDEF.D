// FileName:	ZIPDEF.D
// Purpose:	To Load the requisite functions and procedures
//		for CAD Glove zippers
//		
// Version:	1.01
// Author:		Gary George
// Date:		21.Sep.95
//
// General:
//
   SetData ("EqualityTolerance", 1.e-5) ;
   
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
		 
// Variables:
// ------------------------------------------ Start -------------------------------------------
XY	
	xyInt; 	// if it exist, intersection of line and line
NUMBER 
	nK1, nK2, nM1, nM2, nSlope1, nSlope2, nCase, 
	nY, nX;
// ------------------------------------------ End   -------------------------------------------

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


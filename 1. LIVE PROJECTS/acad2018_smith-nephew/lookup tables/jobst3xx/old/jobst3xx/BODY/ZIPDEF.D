// FileName:	ZIPDEF.D
// Purpose:	To Load the requisite functions and procedures
//		for Bosy suit zippers
//		
// Version:	1.01
// Author:		Gary George
// Date:		21.Sep.95
//
//
//
   
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
   nn = 1; 
   GetVertex ( %1, nn, &xyTmp1) ; 
   while (nn <= nVertex) {
   	nn = nn + 1;
   	xyTmp = xyTmp1 ;
   	GetVertex ( %1, nn, &xyTmp1) ; 
	nLengthToInt = nLengthToInt  + Calc ("length", xyTmp, xyTmp1) ;
	if ( %2 <= nLengthToInt) {
		aAngleAtInt = Calc ("angle", xyTmp, xyTmp1) ;
		xyInt = CalcXY ("relpolar", xyTmp, nLengthToInt - %2, aAngleAtInt ) ;
		return(%true) ;	
		}	
	} 	
    return(%false) ;	
   } //End FN_LengthToPolyInterSection 
  
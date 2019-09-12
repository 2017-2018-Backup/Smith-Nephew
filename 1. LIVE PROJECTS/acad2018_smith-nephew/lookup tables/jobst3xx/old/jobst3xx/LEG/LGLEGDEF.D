// File:		WHLG2DEF.D
// Purpose:	To load Leg \ Lower extremity
//		Drawing functions
//		And string functions
// Version:	1.01
// Date:		07.Feb.94
// Author:	Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action


   sSpecificationNO	= "GOP 01-02/03" ;
   sSpecificationDate = "26 September 91" ;

   nSeam	= 0.1875;

// Functions and procedures to manipiulate strings as Arrays
//


// Function
//	nTapeLength = FNGetTape( nIndex) 
// Returns:
//	nTapeLength at nIndex
//	NB. Converts stored value to a decimal value
// Arguments:
//	nIndex	- Index to nth Tape Length value (Start at 1)
// Globals:
//	sTapeLengths 	- String containing Length of each tape


function FNGetTape (){
  nInt = Value("scalar",StringMiddle(sTapeLengths,( (%1- 1)*4)+1,3));
  nDec =  Value("scalar",StringMiddle(sTapeLengths,( (%1- 1)*4)+4,1));
  return (nInt + (nDec * 0.1)) ;
} /* End FNGetTape */



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
//	nInt		- The value before the decimal point
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

procedure PROpenTemplateFile () {
   // Load template data file
   if (nFabricClass == 0)
   	sFile = sPathJOBST + "\\TEMPLTS\\WH" + StringMiddle(sPressure, 1, 2) + "MMHG.DAT" ;
   else
   	sFile = sPathJOBST + "\\TEMPLTS\\WH" + StringMiddle(sPressure, 1, 2) + "DS.DAT" ;

   hChan = Open ("file", sFile, "readonly") ;
   if (hChan) {
	if (GetLine(hChan, &sLine)){
		ScanLine( sLine, "blank", &nNo, &sScale, &nSpace, &n20Len, &nReduction);
		}
	else
		Exit (%abort, "Can't read " + sFile + "\nFile maybe corrupted");
	}
   else{
	Exit (%abort, "Can't open "+ sFile + "\nCheck installation");
	}
}

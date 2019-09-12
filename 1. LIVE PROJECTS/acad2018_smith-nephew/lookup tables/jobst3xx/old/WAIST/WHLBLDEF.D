// File:		WHLBLDEF.D
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


// Functions and procedures to manipiulate strings as Arrays
//

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
//

function FNDecimalise() {
if (StringCompare("inches", sUnits)) {
	nInt = int(%1);
	nDec = ( %1 - nInt) ; 
	return (nInt + (nDec * 1.25)) ;
	}
else
	return ( %1) ;

} /* End FNDecimalise */

// Layer setting procedures
//

procedure PRSetLayerNotes () {
   hLayer = Table("find", "layer", "Notes") ;
   if ( hLayer > %zero )
 	Execute ("menu", "SetLayer",  hLayer) ;
   else
   	Table ("add", "layer", sLayer, "Notes in Green", "current", Table("find", "color", "Green"));
}



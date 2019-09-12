// File:		WHLEGDEF.D
// Purpose:	Loads  functions
//		To allow drawing of WH legs that use
//		JOBSTEX fabric figured at each tape.
// Version:	1.01
// Date:		02.Dec.94
// Author:	Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action

//  
// Functions and procedures to manipiulate strings as Arrays

// Function:
//	nStretch = FNGetStretch( nIndex) 
// Returns:
//	nStretch
// Arguments:
//	nIndex	- Index to nth sStretch value (Start at 1)
// Globals:
//	sStretch 	- String containing calulated Grams values
//	

function FNGetStretch (){
  return (  Value("scalar",StringMiddle(sStretch,( (%1- 1)*3)+1,3)));
} /* End FNGetStretch */


// Function:
//	nMMs = FNGetMMs( nIndex) 
// Returns:
//	nMMs
// Arguments:
//	nIndex	- Index to nth MMs value (Start at 1)
// Globals:
//	sTapeMMs 	- String containing MM values
//	

function FNGetMMs (){
  return (  Value("scalar",StringMiddle(sTapeMMs,( (%1- 1)*3)+1,3)));
} /* End FNGetMMs */// Function:



// Function
//	nTapeLength = FNGetTapeDisplayFormat( nIndex) 
// Returns:
//	nTapeLength at nIndex
//	Note:-
//	    Returns stored value for display in dialog edit boxes.
//	    Inches and eighths for "inches" units.	
//	    Decimal for "cm" units 
// Arguments:
//	nIndex	- Index to nth Tape Length value (Start at 1)
// Globals:
//	sTapeLengths 	- String containing Length of each tape
//	
// 

function FNGetTapeDisplayFormat (){
  return ( Value("scalar",StringMiddle(sTapeLengths,( (%1- 1)*4)+1,4)) / 10);
} /* End FNGetTapeDisplayFormat */



// Function:
//	nReduction = FNGetReduction( nIndex) 
// Returns:
//	nReduction
// Arguments:
//	nIndex	- Index to nth MMs value (Start at 1)
// Globals:
//	sReduction - String containing calculated reduction values
//	

function FNGetReduction (){
  return (  Value("scalar",StringMiddle(sReduction,( (%1- 1)*3)+1,3)));
} /* End FNGetReduction */





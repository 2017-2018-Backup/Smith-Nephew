// File:		WHPOW.D
// Purpose:	POWERNET  Fabric 
//		Template specific functions
// Version:	1.01
// Date:		04.Dec.94
// Author:	Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
//

//  nHt = FNLegStartHt ( nStartTape , nth Tape) 
//  Uses nStartTape and nLastTape as global variables
//  POWERNET Specific
//
function FNLegStartHt () {
	nn = %1 + 1;
	nLength = 0;
	while (nn <= %2) {
		if (nn <= 4)   nSpace = 1.375;
		if (nn ==  5)  nSpace = 1.21875;
		if (nn ==  6)  nSpace = 1.5;
		if (nn ==  7)  nSpace = 1.5;
		if (nn ==  8)  nSpace = 1.21875;
		if (nn > 8)	    nSpace = 1.25;
		if ((nn == nFirstTape+1) && (nFootPleat1 !=0))      nSpace = nFootPleat1 ;
		if ((nn == nFirstTape+2) && (nFootPleat2 !=0))      nSpace = nFootPleat2 ;
		if ((nn == nLastTape-1)  && (nTopLegPleat2 !=0)) nSpace = nTopLegPleat2 ;
		if ((nn == nLastTape)     && (nTopLegPleat1 !=0) && !FootLess ) nSpace =  nTopLegPleat1 ;
		nLength = nLength + nSpace;
		nn = nn +1 ;
	}
	return (nLength);
} /* End FNLegStartHt */


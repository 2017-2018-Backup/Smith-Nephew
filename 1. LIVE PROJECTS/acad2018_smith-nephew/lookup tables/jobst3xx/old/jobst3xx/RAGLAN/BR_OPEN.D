// File Name:	BR_OPEN.D
// Purpose:	Draw Open or Lining Axilla for a Bodysuit arm
//
// Version:	1.01
// Date:		1.Nov.94
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date	By	Action
// 1.Oct.97   GG	Body suit mods
//--------------------------------------------------------------------------------------------------------------------------------------
// --------------------------------------------------   Open and Lining Axilla -------------------------------------------
//--------------------------------------------------------------------------------------------------------------------------------------
// From the Axilla point traverse the BODYCURV picking up the following intersections
// in the order given (ie Increasing X)
// 
//	1.  xyRaglanNeck,
//	2.  xyBackNeck
//
// For Open Axilla
// 	1.  Establish profile length from xyAxillaBodySuit to xyBackNeck 
//	     find 1/3 Rd distance and
//              2.  Draw arrow and Stamp as open axilla.
//
// For Lining at Axilla
//	1. Stamp Lining

	nLengthToOpenAxillaPt = 0.0 ;
	hCurve = Open ("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly");
   	GetLine(hCurve, &sLine);
   	ScanLine( sLine, "blank", &nLength, &aAngle);
   	aPrevAngle =  -aUpperCurve;   		          // Rotate curve to correct start angle 
    	xyPt1 = xyAxillaBodySuit;
   	DrawSegment = %true ;					// Set draw segments flag on
	StartPoly("polyline");
	AddVertex(xyPt1) ;
	nRadius = nAxillaFrontNeckRad - nFrontNeckRedFac ;
	xyCen = xyAxillaBodySuit;
	nCount = 1;
   	while (GetLine(hCurve, &sLine)){
		aAngle = aAngle + aPrevAngle;
		xyPt2 = CalcXY("relpolar", xyPt1,nLength, aAngle);
		if ( FN_CirLinInt( xyPt1, xyPt2, xyCen , nRadius)) {
			if (nCount == 1) {
				xyRaglanNeck = xyInt;	// Found Raglan end
				nRadius = nAxillaBackNeckRad;
				// Check if it ends on same segment
				if ( FN_CirLinInt( xyPt1, xyPt2, xyCen, nRadius)) nCount = 2 ;
				}
			if (nCount == 2){
				nLengthToOpenAxillaPt = nLengthToOpenAxillaPt + Calc( "length", xyPt1, xyInt) ;
				AddVertex(xyInt) ;
				EndPoly();
			  	// Revise xyEndBottomCurve 
				aAngle = Calc ("angle", xyPt1, xyInt) - 90 ;
				xyEndBottomCurve = CalcXY("relpolar", xyInt, 0.75, aAngle);
				break;
				}
			nCount = nCount + 1;
			}
		if ( DrawSegment )	AddVertex(xyPt2) ;
		nLengthToOpenAxillaPt = nLengthToOpenAxillaPt + Calc( "length", xyPt1, xyPt2) ;
		xyPt1 = xyPt2 ;
		aPrevAngle = aAngle;
		ScanLine( sLine, "blank", &nLength, &aAngle);
   	} //End GetLine while
  	Close ("file", hCurve);
  	
	// Draw Closing line at Axilla
  	AddEntity ("line", xyAxilla, xyAxillaBodySuit);

	PRSetLayerNotes() ;

          if (StringCompare( sAxillaType, "Open")) {
     	// Locate 1/3 raglan length and add an arrow at that point
    	// Label axilla as open
	nLengthToOpenAxillaPt = nLengthToOpenAxillaPt / 3 ;
	nTraversedLength = 0.0 ;
	hCurve = Open ("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly");
   	GetLine(hCurve, &sLine);
   	ScanLine( sLine, "blank", &nLength, &aAngle);
   	aPrevAngle = (Calc("angle",xyAxillaBodySuit, xyBackNeck ) - aUpperCurve);    // Rotate curve to correct start angle
    	xyPt1 = xyAxillaBodySuit;
   	while (GetLine(hCurve, &sLine)){
		aAngle = aAngle + aPrevAngle;
		xyPt2 = CalcXY("relpolar", xyPt1,nLength, aAngle);
		nTraversedLength = nTraversedLength + nLength  ;
		if ( nTraversedLength  > nLengthToOpenAxillaPt ) {
			nError = FN_CirLinInt ( xyPt1, xyPt2, xyPt1, nTraversedLength - nLengthToOpenAxillaPt) ;
			AddEntity("marker",
				"medarrow", 
				xyInt, 
				0.2, 
				0.1, 
				Calc("angle", xyPt1, xyPt2) - 90) ;
			// Add text
			AddEntity("text", 
				"OPEN",
				 xyAxillaBodySuit.x,
				 xyAxillaBodySuit.y - 0.5);
			AddEntity("text", 
				"AXILLA",
				 xyAxillaBodySuit.x,
				 xyAxillaBodySuit.y - 0.7);
				break;
				}
			xyPt1 = xyPt2 ;
			aPrevAngle = aAngle;
			ScanLine( sLine, "blank", &nLength, &aAngle);
   		} //End GetLine while
  		Close ("file", hCurve);
                }
        else { 
		AddEntity("text", 
		"LINING",
		xyAxillaBodySuit.x,
		xyAxillaBodySuit.y - 0.5);
	}
	

// File:	BODOTHRS.D
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
// 17.Jun.94	gg	MODS JUNE.94 (1)

//--------------------------------------------------------------------------------------------------------------------------------------
// --------------------------------------------------   Open and Lining Axilla ------------------------------------------
//--------------------------------------------------------------------------------------------------------------------------------------
// From the Axilla point traverse the VESTCURV picking up the following intersections
// in the order given (ie Increasing X)
// 
//	1.  xyRaglanNeck,
//
// For Open Axilla
// 	1.  Establish profile length from xyAxilla to xyRaglanNeck 
//	     find 1/3 Rd distance and
//              2.  Draw arrow and Stamp as open axilla.
//
// For Lining at Axilla
//	1. Stamp Lining

   if (StringCompare( sAxillaType, "Open") || StringCompare( sAxillaType, "Lining") ){ 
	xyAxillaLow = xyAxilla ;
	nLengthToOpenAxillaPt = 0.0 ;
	hCurve = Open ("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly");
   	GetLine(hCurve, &sLine);
   	ScanLine( sLine, "blank", &nLength, &aAngle);
	aPrevAngle = (aVestCurve - aCurve); 
    	xyPt1 = xyAxilla;
   	DrawSegment = %true ;					// Set draw segments flag on
   	StartPoly("polyline");
   	while (GetLine(hCurve, &sLine)){
		aAngle = aAngle + aPrevAngle;
		xyPt2 = CalcXY("relpolar", xyPt1,nLength, aAngle);
		if ( FN_CirLinInt( xyPt1, xyPt2, xyBackNeckCen,  nBackNeckRadius )) {
//			AddEntity("line", xyPt1, xyInt);
			AddVertex( xyPt1);
			nLengthToOpenAxillaPt = nLengthToOpenAxillaPt + Calc( "length", xyPt1, xyInt) ;
			xyRaglanNeck = xyInt;			// Found Raglan end
			AddVertex( xyRaglanNeck);
			EndPoly();
          		hEnt = UID ("find", UID ("getmax")) ;
    			SetDBData(hEnt, "Data",sID);
    			SetDBData(hEnt, "curvetype","vest" + sSide + "raglan");
			break;
			}
                if ( DrawSegment ) AddVertex(xyPt1);  // Draw raglan
		nLengthToOpenAxillaPt = nLengthToOpenAxillaPt + Calc( "length", xyPt1, xyPt2) ;
		xyPt1 = xyPt2 ;
		aPrevAngle = aAngle;
		ScanLine( sLine, "blank", &nLength, &aAngle);
   	} //End GetLine while
          Close ("file", hCurve);
          PRSetLayerNotes() ;

          if (StringCompare( sAxillaType, "Open")) {
     	// Locate 1/3 raglan length and add an arrow at that point
    	// Label axilla as open
	nLengthToOpenAxillaPt = nLengthToOpenAxillaPt / 3 ;
	nTraversedLength = 0.0 ;
	hCurve = Open ("file", sPathJOBST +  "\\TEMPLTS\\VESTCURV.DAT", "readonly");
   	GetLine(hCurve, &sLine);
   	ScanLine( sLine, "blank", &nLength, &aAngle);
	aPrevAngle = (aVestCurve - aCurve); 
     	xyPt1 = xyAxilla;
   	while (GetLine(hCurve, &sLine)){
		aAngle = aAngle + aPrevAngle;
		xyPt2 = CalcXY("relpolar", xyPt1,nLength, aAngle);
		nTraversedLength = nTraversedLength + nLength  ;
		if ( nTraversedLength  > nLengthToOpenAxillaPt ) {
			nError = FN_CirLinInt ( xyPt1, xyPt2, xyPt1, nTraversedLength - nLengthToOpenAxillaPt) ;
			hEnt = AddEntity("marker",
				"medarrow", 
				xyInt, 
				0.2, 
				0.1, 
				Calc("angle", xyPt1, xyPt2) - 90) ;
    			SetDBData(hEnt, "Data",sID);
    			SetDBData(hEnt, "curvetype","vest" + sSide + "openaxillamarker");
			// Add text
			AddEntity("text", 
				"OPEN",
				 xyAxilla.x - 0.25,
				 xyAxilla.y - 0.5);
			AddEntity("text", 
				"AXILLA",
				 xyAxilla.x - 0.25,
				 xyAxilla.y - 0.7);
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
	xyAxilla.x - 0.25,
	xyAxilla.y - 0.5);
	}
	if ( StringCompare(sSide, "Right")) PRSetLayerTemplateRight () ;
		else PRSetLayerTemplateLeft () ;	
   } // Endif

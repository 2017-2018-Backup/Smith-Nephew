//File 	BODREGUL.D
//	REGULAR AXILLA
//	~~~~~~~~~~~~~~~~~~~~~~~
// From the Axilla point traverse the VESTCURV picking up the following intersections
// in the order given (ie Increasing X)
// 
//	1.  Start of insert allowing for a seam
// 	2.  A tempory construction point 
//	3.  xyRaglanAxilla,
//	4.  xyRaglanNeck,
//

if (StringCompare("Regular", sAxillaType, 7)){ 

   if (StringCompare(  "Regular 2\"",sAxillaType)) nRegAxillaFac_1 = nRegAxillaNormal ;
   if (StringCompare(  "Regular 1½\"",sAxillaType)) nRegAxillaFac_1 = nRegAxillaChild ;
   if (StringCompare(  "Regular 2½\"",sAxillaType)) nRegAxillaFac_1 = nRegAxillaLargeAdult ;
   if (BRAGiven && FN_CirLinInt(xyBreast, xyAxilla, xyAxilla, nRegAxillaFac_1) )
	xyAxillaLow = xyInt;
   else if (FN_CirLinInt(xyWaistOFF, xyAxilla, xyAxilla, nRegAxillaFac_1))
		xyAxillaLow = xyInt;
	else 
		Display ("message", "error", "Can't form Axilla with this data\n");

   xyCen = xyAxilla;				
   nRadius = nSeamAllowance ;
   nCount = 1;
   if (StringCompare(sSide, "Right")) {
	nXInsert =  -10 ;	// Right
	nYInsert =  0 ;
	}
   else {
	nXInsert = -10 ;	// Left
	nYInsert =  3 ;
	}
	
// Open Curve for Reading
// Assume that if the tests for Existance/Corruption are passed. Then no need to repeat them

   hCurve = Open ("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly");
   GetLine(hCurve, &sLine);
   ScanLine( sLine, "blank", &nLength, &aAngle);

   aPrevAngle = (aVestCurve - aCurve);      // Rotate curve to correct start angle

   xyPt1 = xyAxilla;
   DrawSegment = %false ;		// Set draw segments flag of
   DrawInsert = %false ;		//
   StartPoly("polyline");
   
   while (GetLine(hCurve, &sLine)){
	aAngle = aAngle + aPrevAngle;
	xyPt2 = CalcXY("relpolar", xyPt1,nLength, aAngle);

	if ( FN_CirLinInt( xyPt1, xyPt2, xyCen, nRadius)) {
		if (nCount == 1) {
			DrawInsert = %true ;		// Draw insert from seam -
			xyPt1 = xyInt ;
			xyInsertConstruct_3 = xyInt ;
			nRadius = nRegAxillaFac_1;		//   allowance
			}
		if (nCount == 2) {
			xyCen = xyInt;			// Found Construction point
			nRadius = nRegAxillaFac_2;
			}
		if (nCount == 3) {
			xyCen = xyBackNeckCen;		// Found raglan Start
			nRadius = nBackNeckRadius;
			xyRaglanAxilla = xyInt;
			hEnt = AddEntity("line", TransXY (xyPt1, "trs",  nXInsert, nYInsert),
				           TransXY (xyInt, "trs",  nXInsert, nYInsert)); // Draw insert
    			SetDBData(hEnt, "Data",sID);
			xyPt1 = xyInt ;
			// Check if xyRaglanAxilla is too close to frontneck
			if (Calc("length",xyRaglanAxilla, xyFrontNeckOFF) < nInsertToNeckMinDist) {
				Display ("message", "error",
						 "WARNING, Axilla Insert line is too close to Front Neck\nTry a smaller Axilla insert\nModule Cancelled");
				Close("file",hCurve);
				EndPoly();
				Exit (%cancel, "Module Aborted");
				}
			DrawSegment = %true;		// Start Drawing from here
			DrawInsert = %false;		// Stop drawing insert
			}
		if (nCount == 4) {
			xyRaglanNeck = xyInt;		// Found Raglan end
//			AddEntity("line", xyPt1, xyRaglanNeck);
			AddVertex( xyPt1);
			AddVertex( xyRaglanNeck);
			EndPoly();
          		hEnt = UID ("find", UID ("getmax")) ;
    			SetDBData(hEnt, "Data",sID);
    			SetDBData(hEnt, "curvetype","vest" + sSide + "raglan");
			break;
			}
		nCount = nCount +1 ;
		}

//         if ( DrawSegment ) AddEntity("line",xyPt1,xyPt2);  // Draw raglan
          if ( DrawSegment ) AddVertex (xyPt1);  // Draw raglan
	if ( DrawInsert )    	AddEntity("line", TransXY (xyPt1, "trs",  nXInsert, nYInsert),
			                          TransXY (xyPt2, "trs",  nXInsert, nYInsert)); // Draw insert
	xyPt1 = xyPt2;
	aPrevAngle = aAngle;
	ScanLine( sLine, "blank", &nLength, &aAngle);
   } //End GetLine while

   Close ("file", hCurve);

// Regular Axilla Construction point
//
   aAngle = acos ( nRegAxillaFac_1/ Calc("length",xyAxilla, xyRaglanAxilla));
   aAngle = Calc("angle",xyAxilla, xyRaglanAxilla) - aAngle;
   xyAxillaConstruct_2 = CalcXY("relpolar", xyAxilla, nRegAxillaFac_1, aAngle);
   xyInsertConstruct_1 = CalcXY("relpolar", xyAxilla, nRegAxillaFac_1 + nSeamAllowance, aAngle);
   xyInsertConstruct_2 = CalcXY("relpolar", xyRaglanAxilla, nSeamAllowance, aAngle);
   } // End if for Regular axilla

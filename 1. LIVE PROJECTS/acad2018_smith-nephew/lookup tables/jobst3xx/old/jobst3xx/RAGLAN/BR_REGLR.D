// File Name:	BR_REGLR.D
// Purpose:		Draw Regular axilla 
//
// Version:		1.01
// Date:		1.Nov.94
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date	By	Action
// 01.Oct.97	GG	Revisons wrt body suit

//--------------------------------------------------------------------------------------------------------------------------------------
// --------------------------------------------------   Regular Axilla --------------------------------------------------------
//--------------------------------------------------------------------------------------------------------------------------------------
// From the Axilla point traverse the BODYCURV picking up the following intersections
// in the order given (ie Increasing X)
// 
//	1.  Start of insert allowing for a seam
// 	2.  A tempory construction point 
//	3.  xyRaglanAxilla,
//	4.  xyRaglanNeck,
//
//

   	if (StringCompare(  "Regular 2\"",sAxillaType))
		nRegAxillaFac_1 = nRegAxillaNormal ;
	if (StringCompare(  "Regular 1½\"",sAxillaType))
		nRegAxillaFac_1 = nRegAxillaChild ;
	if (StringCompare(  "Regular 2½\"",sAxillaType))
		nRegAxillaFac_1 = nRegAxillaLargeAdult ;

	nXInsert = -10 ;
	nYInsert = 3.5;

  	xyCen = xyAxillaBodySuit;				

	// Open Curve for Reading
   	hCurve = Open ("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly");
   	GetLine(hCurve, &sLine);
   	ScanLine( sLine, "blank", &nLength, &aAngle);

//   	aPrevAngle = (Calc("angle",xyAxilla, xyBackNeck ) - aUpperCurve);      // Rotate curve to correct start angle
   	aPrevAngle =  -aUpperCurve;      // Rotate curve to correct start angle

   	xyPt1 = xyAxillaBodySuit;
   	DrawSegment = %false ;		// Set draw segments flag off
	DrawInsert = %false;
  	nRadius = nSeamAllowance;
  	nCount = 1;
	   while (GetLine(hCurve, &sLine)){
		aAngle = aAngle + aPrevAngle;
		xyPt2 = CalcXY("relpolar", xyPt1,nLength, aAngle);
		if ( FN_CirLinInt( xyPt1, xyPt2, xyCen, nRadius)) {
			if (nCount == 1) {
				DrawInsert = %true;
				xyPt1 = xyInt;			// Found Construction point
				xyInsertConstruct_3 = xyInt;
				nRadius = nRegAxillaFac_1;
				}
			if (nCount == 2) {
				xyCen = xyInt;			// Found Construction point
				nRadius = nRegAxillaFac_2;
				}
			if (nCount == 3) {
				xyCen = xyAxillaBodySuit;			// Found raglan Start
				nRadius =  nAxillaFrontNeckRad - nFrontNeckRedFac ;
				xyRaglanAxilla = xyInt;
				AddEntity("line", TransXY (xyPt1, "trs",  nXInsert, nYInsert),
				           TransXY (xyInt, "trs",  nXInsert, nYInsert)); // Draw insert
				xyPt1 = xyInt ;
				DrawSegment = %true;		// Start Drawing from here
				StartPoly("polyline");
				AddVertex(xyPt1);
				DrawInsert = %false;
				}
			if (nCount == 4) {
				xyRaglanNeck = xyInt;		// Found Raglan end
				xyCen    = xyAxillaBodySuit;
				nRadius = nAxillaBackNeckRad;
				// Check if it ends on same segment
				if ( FN_CirLinInt( xyPt1, xyPt2, xyCen, nRadius)) nCount = 5 ;
				}
			if (nCount == 5){
//				AddVertex(xyBackNeck);
				AddVertex(xyInt);
				EndPoly();
			  	// Revise xyEndBottomCurve 
				aAngle = Calc ("angle", xyPt1,xyInt) - 90 ;
				xyEndBottomCurve = CalcXY("relpolar",xyInt, 0.75, aAngle);
				break;
				}
			nCount = nCount +1 ;
			}
	         	       if ( DrawSegment ) AddVertex(xyPt2);
		       if ( DrawInsert )      AddEntity("line", TransXY (xyPt1, "trs",  nXInsert, nYInsert)
			                                              ,  TransXY (xyPt2, "trs",  nXInsert, nYInsert)); // Draw insert
		xyPt1 = xyPt2 ;
		aPrevAngle = aAngle;
		ScanLine( sLine, "blank", &nLength, &aAngle);
  	 } //End GetLine while

    Close ("file", hCurve);

    // Draw insert 
    // Calculate second Regular Axilla Construction point and insert points
    //
    aAngle = acos ( nRegAxillaFac_1/ Calc("length",xyAxillaBodySuit, xyRaglanAxilla));
    aAngle = Calc("angle",xyAxillaBodySuit, xyRaglanAxilla) - aAngle;
    xyAxillaConstruct_2 = CalcXY("relpolar", xyAxillaBodySuit, nRegAxillaFac_1, aAngle);
    xyInsertConstruct_1 = CalcXY("relpolar", xyAxillaBodySuit, nRegAxillaFac_1 + nSeamAllowance, aAngle);
    xyInsertConstruct_2 = CalcXY("relpolar", xyRaglanAxilla, nSeamAllowance, aAngle);

    // Find point on sleeve profile (note - this is an aproximation only)
    // From above nTape is currenty set to last vertex on the profile
    // For curve profiles only, not short sleeves 3 tapes or less
      if (StringCompare("curve", sEntClass) || StringCompare("polyline", sEntClass)) {
    	nTape = nTape - 1 ;  
    	xyPt1 = xyAxilla ;
   	while (nTape > 0 ) {
 		GetVertex( hSleeveProfile, nTape, &xyPt2) ;
		if (FN_CirLinInt ( xyPt1, xyPt2, xyAxillaBodySuit,  nRegAxillaFac_1)) {
			xyAxillaLow = xyInt ;
    			hEnt = AddEntity ( "line", xyAxillaConstruct_2, xyRaglanAxilla);
			SetDBData(hEnt, "ID", sID + "RegularAxillaLine");
        			aPrevAngle = Calc ( "angle", xyAxillaBodySuit, xyAxillaLow);
   			aAngle =  Calc ( "angle",  xyAxillaBodySuit, xyAxillaConstruct_2) - aPrevAngle ;
   			hEnt = AddEntity ( "arc", xyAxillaBodySuit, nRegAxillaFac_1, aPrevAngle, aAngle); 
			SetDBData(hEnt, "ID", sID + "RegularAxillaArc");
			break ;	
			}
    	 	nTape = nTape - 1;	
		xyPt1 = xyPt2 ;
		}
	}
    else  {
	if (FN_CirLinInt ( xyProfileStart, xyAxilla, xyAxilla,  nRegAxillaFac_1)) {
		xyAxillaLow = xyInt ;
    		hEnt = AddEntity ( "line", xyAxillaConstruct_2, xyRaglanAxilla);
    		SetDBData(hEnt, "ID", sID + "RegularAxillaLine") ;
    		aPrevAngle = Calc ( "angle", xyAxillaBodySuit, xyAxillaLow);
   		aAngle =  Calc ( "angle",  xyAxillaBodySuit, xyAxillaConstruct_2) - aPrevAngle ;
   		hEnt = AddEntity ( "arc", xyAxillaBodySuit, nRegAxillaFac_1, aPrevAngle, aAngle); 
    		SetDBData(hEnt, "ID", sID + "RegularAxillaArc") ;
    		// Warn if too close to end of sleeve
		if (Calc("length", xyAxillaLow, xyProfileStart) < 0.5)
			Display("message", "error", "Axilla Insert is too close to end of Sleeve, need minimum of 1/2\"");
//			Exit (%cancel, "Could not make an Axilla Insert\nToo close to end of Sleeve, need minimum of 3/4\"") ;
		}
	else
		Exit (%cancel, "Could not make an Axilla Insert\nTry again with a smaller Axilla Insert size") ;
	}
	
// Warn if too close to elbow
//
   if (Calc("length", xyAxillaLow, xyElbow) < 0.75)
	Display("message", "error", "Axilla Insert is too close to the Elbow, need minimum of 3/4\"");

// Draw insert
//
   xyInsertConstruct_4 = CalcXY ("relpolar", xyAxillaLow
				 , sqrt ( 2*(nSeamAllowance*nSeamAllowance))
				 , aPrevAngle + 45) ;
   aPrevAngle = Calc ( "angle", xyAxillaBodySuit, xyInsertConstruct_4  );
   aAngle =  Calc ( "angle",  xyAxillaBodySuit, xyAxillaConstruct_2) - aPrevAngle ;

   hEnt = AddEntity ( "arc", TransXY( xyAxillaBodySuit, "trs", nXInsert, nYInsert)
	          , nRegAxillaFac_1 + nSeamAllowance, aPrevAngle, aAngle); 
   SetDBData(hEnt, "ID", sID + "RegularAxillaInsertArc") ;
   hEnt = AddEntity ("line", TransXY(xyInsertConstruct_4, "trs", nXInsert, nYInsert)
	          , TransXY(xyInsertConstruct_3, "trs", nXInsert, nYInsert)) ;
   SetDBData(hEnt, "ID", sID + "RegularAxillaInsertLine") ;
   AddEntity ("line", TransXY(xyInsertConstruct_1, "trs", nXInsert, nYInsert)
	          , TransXY(xyInsertConstruct_2, "trs", nXInsert, nYInsert)) ;	
   AddEntity ("line", TransXY(xyInsertConstruct_2, "trs", nXInsert, nYInsert)
	          , TransXY(xyRaglanAxilla, "trs", nXInsert, nYInsert)) ;	

   // Draw seam line and stamp 
   // Note recalculation because of drawing on the notes layer
   PRSetLayerNotes();
   AddEntity ( "line", TransXY(xyAxillaConstruct_2, "trs", nXInsert, nYInsert)
	           , TransXY(xyRaglanAxilla, "trs", nXInsert, nYInsert)); 
   aPrevAngle = Calc ( "angle", xyAxilla,xyInsertConstruct_4);
   aAngle =  Calc ( "angle",  xyAxillaBodySuit, xyAxillaConstruct_2) - aPrevAngle ;
   hEnt = AddEntity ( "arc", TransXY( xyAxillaBodySuit, "trs", nXInsert, nYInsert)
	          , nRegAxillaFac_1, aPrevAngle, aAngle); 
   SetDBData(hEnt, "ID", sID + "RegularAxillaInsertSeamArc") ;
   if (StringCompare("Left",sSleeve)) 
   	PRDataStamp ( TransXY(xyAxillaBodySuit.x - 1, xyAxillaBodySuit.y - 1, "trs", nXInsert, nYInsert), 21) ;
   else
   	PRDataStamp ( TransXY(xyAxillaBodySuit.x - 1, xyAxillaBodySuit.y - 1, "trs", nXInsert, nYInsert), 22) ;

   // End  Regular Axilla


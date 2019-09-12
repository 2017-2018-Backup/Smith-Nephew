// File Name:	SR_REGLR.D
// Purpose:	Draw Regular axilla
//
// Version:	1.01
// Date:		1.Nov.94
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date	By	Action

//--------------------------------------------------------------------------------------------------------------------------------------
// --------------------------------------------------   Regular Axilla --------------------------------------------------------
//--------------------------------------------------------------------------------------------------------------------------------------
// From the Axilla point traverse the VESTCURV picking up the following intersections
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

  	xyCen = xyAxilla;				

	// Open Curve for Reading
   	hCurve = Open ("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly");
   	GetLine(hCurve, &sLine);
   	ScanLine( sLine, "blank", &nLength, &aAngle);

   	aPrevAngle = (Calc("angle",xyAxilla, xyBackNeck ) - aUpperCurve);      // Rotate curve to correct start angle

   	xyPt1 = xyAxilla;
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
				xyCen = xyAxilla;			// Found raglan Start
				nRadius =  nAxillaFrontNeckRad - nFrontNeckRedFac ;
				xyRaglanAxilla = xyInt;
				hEnt = AddEntity("line", TransXY (xyPt1, "trs",  nXInsert, nYInsert),
				           TransXY (xyInt, "trs",  nXInsert, nYInsert)); // Draw insert
				xyPt1 = xyInt ;
  				SetDBData(hEnt, "ID",sID);
    				SetDBData(hEnt, "curvetype","sleeveinsert");
    				SetDBData(hEnt, "Data", sData);
				DrawSegment = %true;		// Start Drawing from here
				StartPoly("polyline");
				AddVertex(xyPt1);
				DrawInsert = %false;
				}
			if (nCount == 4) {
				xyRaglanNeck = xyInt;		// Found Raglan end
				xyCen    = xyAxilla;
				nRadius = nAxillaBackNeckRad;
				// Check if it ends on same segment
				if ( FN_CirLinInt( xyPt1, xyPt2, xyCen, nRadius)) nCount = 5 ;
				}
			if (nCount == 5){
				AddVertex(xyBackNeck);
				EndPoly();
          			hEnt = UID ("find", UID ("getmax")) ;
    				SetDBData(hEnt, "curvetype","sleeveraglan");
    				SetDBData(hEnt, "Data", sData);
				break;
				}
			nCount = nCount +1 ;
			}
	         	       if ( DrawSegment ) AddVertex(xyPt2);
		       if ( DrawInsert )   {   	hEnt = AddEntity("line", TransXY (xyPt1, "trs",  nXInsert, nYInsert)
			                                              ,  TransXY (xyPt2, "trs",  nXInsert, nYInsert)); // Draw insert
  						SetDBData(hEnt, "ID",sID);
    						SetDBData(hEnt, "curvetype","sleeveinsert");
    						SetDBData(hEnt, "Data", sData);			                             
			                         }
		xyPt1 = xyPt2 ;
		aPrevAngle = aAngle;
		ScanLine( sLine, "blank", &nLength, &aAngle);
  	 } //End GetLine while

    Close ("file", hCurve);

    // Draw insert 
    // Calculate second Regular Axilla Construction point and insert points
    //
    aAngle = acos ( nRegAxillaFac_1/ Calc("length",xyAxilla, xyRaglanAxilla));
    aAngle = Calc("angle",xyAxilla, xyRaglanAxilla) - aAngle;
    xyAxillaConstruct_2 = CalcXY("relpolar", xyAxilla, nRegAxillaFac_1, aAngle);
    xyInsertConstruct_1 = CalcXY("relpolar", xyAxilla, nRegAxillaFac_1 + nSeamAllowance, aAngle);
    xyInsertConstruct_2 = CalcXY("relpolar", xyRaglanAxilla, nSeamAllowance, aAngle);

    // Find point on sleeve profile (note - this is an aproximation only)
    // From above nTape is currenty set to last vertex on the profile
    // For curve profiles only, not short sleeves 3 tapes or less
      if (StringCompare("curve", sEntClass) || StringCompare("polyline", sEntClass)) {
    	nTape = nTape - 1 ;  
    	xyPt1 = xyAxilla ;
   	while (nTape > 0 ) {
 		GetVertex( hSleeveProfile, nTape, &xyPt2) ;
		if (FN_CirLinInt ( xyPt1, xyPt2, xyAxilla,  nRegAxillaFac_1)) {
			xyAxillaLow = xyInt ;
    			hEnt = AddEntity ( "line", xyAxillaConstruct_2, xyRaglanAxilla);
			SetDBData(hEnt, "ID", sID + "RegularAxillaLine");
   			SetDBData(hEnt, "Data", sData);  SetDBData(hEnt, "curvetype","sleevecutout");
        			aPrevAngle = Calc ( "angle", xyAxilla, xyAxillaLow);
   			aAngle =  Calc ( "angle",  xyAxilla, xyAxillaConstruct_2) - aPrevAngle ;
   			hEnt = AddEntity ( "arc", xyAxilla, nRegAxillaFac_1, aPrevAngle, aAngle); 
			SetDBData(hEnt, "ID", sID + "RegularAxillaArc");
   			SetDBData(hEnt, "Data", sData);  SetDBData(hEnt, "curvetype","sleevecutout");
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
   		SetDBData(hEnt, "Data", sData);  SetDBData(hEnt, "curvetype","sleevecutout");
    		aPrevAngle = Calc ( "angle", xyAxilla, xyAxillaLow);
   		aAngle =  Calc ( "angle",  xyAxilla, xyAxillaConstruct_2) - aPrevAngle ;
   		hEnt = AddEntity ( "arc", xyAxilla, nRegAxillaFac_1, aPrevAngle, aAngle); 
    		SetDBData(hEnt, "ID", sID + "RegularAxillaArc") ;
   		SetDBData(hEnt, "Data", sData);  SetDBData(hEnt, "curvetype","sleevecutout");
    		SetDBData(hEnt, "Data", sData);			                             
		if (Calc("length", xyAxillaLow, xyProfileStart) < 0.75)
			Exit (%cancel, "Could not make an Axilla Insert\nToo close to end of Sleeve, need minimum of 3/4\"") ;
		}
	else
		Exit (%cancel, "Could not make an Axilla Insert\nTry again with a smaller Axilla Insert size") ;
	}
	

   // Draw insert
   xyInsertConstruct_4 = CalcXY ("relpolar", xyAxillaLow
				 , sqrt ( 2*(nSeamAllowance*nSeamAllowance))
				 , aPrevAngle + 45) ;
   aPrevAngle = Calc ( "angle", xyAxilla, xyInsertConstruct_4  );
   aAngle =  Calc ( "angle",  xyAxilla, xyAxillaConstruct_2) - aPrevAngle ;

   hEnt = AddEntity ( "arc", TransXY( xyAxilla, "trs", nXInsert, nYInsert)
	          , nRegAxillaFac_1 + nSeamAllowance, aPrevAngle, aAngle); 
   SetDBData(hEnt, "ID", sID + "RegularAxillaInsertArc") ;
   SetDBData(hEnt, "Data", sData);  SetDBData(hEnt, "curvetype","sleeveinsert");
   hEnt = AddEntity ("line", TransXY(xyInsertConstruct_4, "trs", nXInsert, nYInsert)
	          , TransXY(xyInsertConstruct_3, "trs", nXInsert, nYInsert)) ;
   SetDBData(hEnt, "ID", sID + "RegularAxillaInsertLine") ;
   SetDBData(hEnt, "Data", sData);  SetDBData(hEnt, "curvetype","sleeveinsert");
   hEnt = AddEntity ("line", TransXY(xyInsertConstruct_1, "trs", nXInsert, nYInsert)
	          , TransXY(xyInsertConstruct_2, "trs", nXInsert, nYInsert)) ;		        
   SetDBData(hEnt, "ID",sID);
   SetDBData(hEnt, "Data", sData);  SetDBData(hEnt, "curvetype","sleeveinsert");
   hEnt = AddEntity ("line", TransXY(xyInsertConstruct_2, "trs", nXInsert, nYInsert)
	          , TransXY(xyRaglanAxilla, "trs", nXInsert, nYInsert)) ;	
   SetDBData(hEnt, "ID",sID);
   SetDBData(hEnt, "Data", sData);  SetDBData(hEnt, "curvetype","sleeveinsert");

   // Draw seam line and stamp 
   // Note recalculation because of drawing on the notes layer
   PRSetLayerNotes();
   hEnt = AddEntity ( "line", TransXY(xyAxillaConstruct_2, "trs", nXInsert, nYInsert)
	           , TransXY(xyRaglanAxilla, "trs", nXInsert, nYInsert)); 
   SetDBData(hEnt, "ID",sID);
   SetDBData(hEnt, "Data", sData);  SetDBData(hEnt, "curvetype","sleeveinsert");
   aPrevAngle = Calc ( "angle", xyAxilla,xyInsertConstruct_4);
   aAngle =  Calc ( "angle",  xyAxilla, xyAxillaConstruct_2) - aPrevAngle ;
   hEnt = AddEntity ( "arc", TransXY( xyAxilla, "trs", nXInsert, nYInsert)
	          , nRegAxillaFac_1, aPrevAngle, aAngle); 
   SetDBData(hEnt, "ID", sID + "RegularAxillaInsertSeamArc") ;
   SetDBData(hEnt, "Data", sData);  SetDBData(hEnt, "curvetype","sleeveinsert");

  // End  Regular Axilla


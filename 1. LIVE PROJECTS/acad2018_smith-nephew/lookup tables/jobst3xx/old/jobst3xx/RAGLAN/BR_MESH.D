// File Name:	BR_MESH.D
// Purpose:		Draw mesh raglan for body suit
//
// Version:		1.01
// Date:		1.Nov.95
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
// 02.Oct.97		GG	Made Body specific
//			  N.B. This is not the final BodySuit mesh as
//				the format of the mesh is still under
//				consideration. 

//--------------------------------------------------------------------------------------------------------------------------------------
// -------------------------------------------------- Mesh Axilla  --------------------------------------------------------
//--------------------------------------------------------------------------------------------------------------------------------------


// From the Axilla point traverse the BODYCURV picking up the following intersections
// in the order given (ie Increasing X)
//
//	1. Start of ragalan curve (1/8th) back from next point
//	2. Intersection of end of Mesh symbol and raglan curve, defines symbol angle
 
   if (nAge <= 14) {
	sMeshSymbol = "mesh0to14";
	nMeshAxillaFac = nMeshAxillaGussetFac ;
	nMeshSymbolLength = nMSLengthGusset;	
	}
   else {
	sMeshSymbol = "mesh15andUP";
	nMeshAxillaFac = nMeshAxillaGussetBoysFac ;
	nMeshSymbolLength = nMSLengthGussetBoys;	
	}

    // Find point on sleeve profile (note - this is an aproximation only)
    // From above nTape is currenty set to last vertex on the profile
    if (StringCompare("curve", sEntClass) || StringCompare("polyline", sEntClass)) {
  	  nTape = nTape - 1 ;  
  	  xyPt1 = xyAxilla ;
  	  while (nTape > 0 ) {
		GetVertex( hSleeveProfile, nTape, &xyPt2) ;
		if (FN_CirLinInt ( xyPt1, xyPt2, xyAxilla,  nMeshAxillaFac - 0.125)) {
			xyAxillaLow = xyInt ;
			}
		// Get seam allowance point
		if (FN_CirLinInt ( xyPt1, xyPt2, xyAxilla,  nMeshAxillaFac )) {
			xyAxillaConstruct_2 = xyInt ;
			break ;	
			}
     		nTape = nTape - 1;	
		xyPt1 = xyPt2 ;
		}
	}
    else {
	if (FN_CirLinInt ( xyProfileStart, xyAxilla, xyAxilla,  nMeshAxillaFac - 0.125)) 
		xyAxillaLow = xyInt ;
	if (FN_CirLinInt (xyProfileStart, xyAxilla, xyAxilla,  nMeshAxillaFac )) 
			xyAxillaConstruct_2 = xyInt ;
		
	}

// Open Curve for Reading
// Assume that if the tests for Existance/Corruption are passed. Then no need to repeat them

   hCurve = Open ("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly");
   GetLine(hCurve, &sLine);
   ScanLine( sLine, "blank", &nLength, &aAngle);

   aPrevAngle = -aUpperCurve ;   
   xyCen = xyAxillaLow;		
   nRadius = nMeshSymbolLength  - nSeamAllowance ;
   nCount = 1;
   xyPt1 = xyAxillaBodySuit;
//   DrawSegment = %false ;		// Set draw segments flag off
   DrawSegment = %true ;		// Draw entire curve for the moment
   StartPoly("polyline");
   AddVertex(xyAxillaBodySuit); 	
   while (GetLine(hCurve, &sLine)){
	aAngle = aAngle + aPrevAngle;
	xyPt2 = CalcXY("relpolar", xyPt1,nLength, aAngle);

	if ( FN_CirLinInt( xyPt1, xyPt2, xyCen, nRadius)) {
		if (nCount == 1) {
			DrawSegment = %true ;		// Draw from seam allowance
			xyRaglanAxilla = xyInt;		
			nRadius = nMeshSymbolLength;	
			// Check if seam allowance on same segment
			if ( FN_CirLinInt( xyPt1, xyPt2, xyCen, nRadius)) nCount = 2 ;
			xyPt1 = xyRaglanAxilla ;
//			StartPoly("polyline");
//			AddVertex(xyRaglanAxilla); 		
			}
		if (nCount == 2) {
			xyMeshAtSeamAllowance = xyInt;	// Mesh symbol end for angle
			xyCen = xyAxillaBodySuit;			
			nRadius =  nAxillaFrontNeckRad - nFrontNeckRedFac ;
			}
		if (nCount == 3) {
			xyRaglanNeck = xyInt;		// Found Raglan end
			xyCen    =  xyAxillaBodySuit;
			nRadius = nAxillaBackNeckRad;
			// Check if it ends on same segment
			if ( FN_CirLinInt( xyPt1, xyPt2, xyCen, nRadius)) nCount = 4 ;
			}
		if (nCount == 4){
			AddVertex(xyBackNeck); 	
			EndPoly();	
			// Revise xyEndBottomCurve 
			aAngle = Calc ("angle", xyPt1, xyInt) - 90 ;
			xyEndBottomCurve = CalcXY("relpolar", xyInt, 0.75, aAngle);
			break;
			}
		nCount = nCount +1 ;
		}

                if ( DrawSegment ) AddVertex(xyPt2); 	  // Draw raglan
	xyPt1 = xyPt2;
	aPrevAngle = aAngle;
	ScanLine( sLine, "blank", &nLength, &aAngle);
   } //End GetLine while

   Close ("file", hCurve);

// Insert Mesh axilla symbol at correct angle 
// Insert a marker to show start point for manual editing
//   AddEntity ("marker", "xmarker", xyAxillaLow,  0.0625,  0.0625, 0);		// Mesh point
    PRSetLayerNotes() ;
//   AddEntity ("marker", "xmarker", xyAxillaConstruct_2, 0.0625,  0.0625, 0);	// Seam allowance
   
//   aAngle = Calc("angle", xyAxillaLow, xyMeshAtSeamAllowance);

//   SetSymbolLibrary( sPathJOBST + "\\JOBST.SLB"); 
//   if ( !Symbol("find", sMeshSymbol+"seam")) Exit(%cancel, "Can't find MESH Axilla symbol to insert\nCheck your installation, that JOBST.SLB exists");
//   if ( !Symbol("find", sMeshSymbol+"pro")) Exit(%cancel, "Can't find MESH Axilla symbol to insert\nCheck your installation, that JOBST.SLB exists");
// Insert  Seam
//   hMesh = AddEntity("symbol", sMeshSymbol + "seam", xyAxillaLow, 1, 1, aAngle);
   SetData("TextHorzJust", 4 );
   SetData("TextVertJust", 32 );
   if (nAge <= 14) 
   	AddEntity("text", 
		"1-3/4\" GUSSET",
	 	xyAxillaLow.x - 0.25,
	 	xyAxillaLow.y - 0.75);
   else 
   	AddEntity("text", 
		"BOYS GUSSET",
	 	xyAxillaLow.x - 0.25,
	 	xyAxillaLow.y - 0.75);
	 	
// Insert profile
//    if (StringCompare(sSleeve, "Right"))
//	PRSetLayerTemplateRight() ;
//    else
//	PRSetLayerTemplateLeft() ;
//    hMesh = AddEntity("symbol", sMeshSymbol + "pro", xyAxillaLow, 1, 1, aAngle);

// Revised code to write a datafile and start the MESHDRAW
// 
   hCurve = Open ("file", "C:\\JOBST\\MESHDRAW.DAT", "write") ;
	SetData("UnitLinearType", 0);	// "Inches"
	PrintFile(hCurve, xyAxillaConstruct_2, "\n") ;
	PrintFile(hCurve, xyAxilla, "\n") ;
	PrintFile(hCurve, xyAxillaBodySuit, "\n") ;
	PrintFile(hCurve, xyBackNeck, "\n") ;
	PrintFile(hCurve, nMeshLength, "\n") ;
	PrintFile(hCurve, nDistanceAlongRaglan, "\n") ;
	PrintFile(hCurve, sSleeve, "\n") ;
	PrintFile(hCurve, sID, "\n") ;
	SetData("UnitLinearType", 6);	// "Inches/Fraction"
   Close ("file", hCurve) ;

// End  MESH Axilla



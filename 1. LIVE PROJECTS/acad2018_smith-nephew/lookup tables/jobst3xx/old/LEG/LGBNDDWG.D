// File Name:	LGBNDDWG.D
// Purpose:	Draw leg for thigh bands and 
//		footless thigh lengths
//
// Version:	1.01
// Date:		15.Dec.94
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action


// Establish layer
   if (StringCompare(sLeg,"Left")) {
	PRSetLayerTemplateLeft () ;	
	hTemplateLayer = Table("find", "layer", "TemplateLeft");
	}
   else	{
	PRSetLayerTemplateRight () ;	
	hTemplateLayer = Table("find", "layer", "TemplateRight");
	}

// draw on layer construct
// 'cause drafix is £"^%%%&*&(*&@@ 
   PRSetLayerConstruct();
	
// Draw Leg 
//
   PROpenTemplateFile ();

// Skip to FirstTape
    nn = 1;
    while (nn < nStyleFirstTape  ) {
	GetLine(hChan, &sLine) ;
	nn = nn + 1;
	} 

   ScanLine( sLine, "blank", &nNo, &sScale, &nSpace, &n20Len, &nReduction);

    xyTmp = xyOtemplate ;

// Start drawing profile
// Note, use of polyline for bands with only 3 tapes
//
    if (nStyleLastTape - nStyleFirstTape == 2) 
	StartPoly("polyline") ;
    else
    	StartPoly ("fitted") ;

    while (nn <= nStyleLastTape ) {
	nTapeLen = FNRound(FNDecimalise(FNGetTape(nn)) * nUnitsFac) ;
	nLength = n20Len / 20 * nTapeLen ;
	nRedStep = n20Len /(20 * 8) ;		
	if (nn == nStyleFirstTape && nStyleFirstRed > 0)
		// Release last tape to a given reduction
		nLength = (nTapeLen * ((100 - nStyleFirstRed ) /100) ) / 2 ;

	if (nn == nStyleLastTape && nStyleLastRed > 0)
		// Release last tape to a given reduction
		nLength = (nTapeLen * ((100 - nStyleLastRed ) /100) ) / 2 ;

	xyTmp.y = xyOtemplate.y + nSeam + nLength ;
	AddVertex (xyTmp) ; 	
	if (nn == nStyleLastTape ) 
		xyProfileLast = xyTmp;
	if (nn == nStyleLastTape - 1 ) 
		xyPrevProfileLast = xyTmp;
	if (nn == nStyleFirstTape ) 
		xyProfileStart = xyTmp;

	nn = nn + 1;
	GetLine(hChan, &sLine) ;
	ScanLine( sLine, "blank", &nNo, &sScale, &nSpace, &n20Len, &nReduction);
	if ((nn == nStyleFirstTape+1) && (nFootPleat1 !=0)) nSpace = nFootPleat1 ;
	if ((nn == nStyleFirstTape+2) && (nFootPleat2 !=0)) nSpace = nFootPleat2 ;
	if ((nn == nStyleLastTape) && (nTopLegPleat1 !=0)) nSpace =  nTopLegPleat1 ;
	if ((nn == nStyleLastTape-1) && (nTopLegPleat2 !=0)) nSpace =  nTopLegPleat2 ;
	xyTmp.x = xyTmp.x + nSpace ;
	}
    Close ("file", hChan) ;

    EndPoly () ;

    // Get polyline entity handle
    // Change layer and set DB values
    if (nStyleLastTape - nStyleFirstTape == 2) 
	sTmp = "layer = 'Construct' AND type = 'Polyline'" ;
    else
    	sTmp = "layer = 'Construct' AND type = 'Curve'" ;

    hChan = Open ("selection", sTmp);
    if( hChan ) {
	ResetSelection (hChan);
	hCurv = GetNextSelection(hChan);
	SetEntityData ( hCurv, "layer", hTemplateLayer);
   	SetDBData(hCurv, "ID", sLegStyle+sFileNo+sLeg+"LegCurve") ;
	}
    Close ("selection", hChan);

    Execute ("menu", "SetLayer", hTemplateLayer);

    if (nLegStyle == 2 || nLegStyle == 4 || nLegStyle == 5 ) {   
	// All nThighPlt values are calculated in LBLEGDIA VB program
	// All nThighPlt values are calculated in LBLEGDIA VB program
	xyO.x = xyO.x + nThighTopExtension ;
	
	// Template end point
	aAngle = Calc("angle", xyPrevProfileLast, xyProfileLast) ;
	nB =  (xyO.x - nThighPltXoff ) - xyPrevProfileLast.x ;
	nA = tan (aAngle) * nB ;

	xyThighPlt.x = xyPrevProfileLast.x + nB;
	xyThighPlt.y = xyPrevProfileLast.y + nA;
	
	// Center of template arc (normal case)
	nB = nThighPltRad - nThighPltXoff ;
	nA = sqrt ( nThighPltRad ^ 2 - nB ^2) ;
	
	xyTmp.x = xyThighPlt.x - nB ; 	// Arc Center X
	xyTmp.y = xyThighPlt.y - nA ;		// Arc Center Y
	
	if ( xyTmp.y < xyO.y) {
		// Special case where center point is below fold line
		xyPt1.y = xyO.y + 0.5 ;
		xyPt1.x = xyO.x ;
		nLength = Calc("length", xyPt1, xyThighPlt) ;
		aAngle = Calc("angle", xyPt1, xyThighPlt) ;
		nA = sqrt ( nThighPltRad ^ 2 - (nLength / 2) ^2) ;
		xyTmp = CalcXY("relpolar", CalcXY("relpolar", xyPt1, nLength / 2, aAngle), nA, aAngle + 90) ;
		nThighPltStartAngle = Calc ("angle", xyTmp, xyPt1) ;	
		nThighPltDeltaAngle = Calc ("angle", xyTmp, xyThighPlt)  - nThighPltStartAngle ;
		}	
	else {
		// Normal case
		nThighPltStartAngle = 0 ;	
		nThighPltDeltaAngle = Calc("angle", xyTmp, xyThighPlt)  - nThighPltStartAngle ;
		}

	hEnt = AddEntity ("arc", xyTmp,
			      nThighPltRad,
		           	      nThighPltStartAngle,
		           	      nThighPltDeltaAngle);	
	SetDBData(hEnt, "ID",  sLegStyle+sFileNo+sLeg+"ClosingLine") ;	

	// Bottom closing line
	hEnt = AddEntity ("line", xyO, CalcXY("relpolar", xyTmp, nThighPltRad, nThighPltStartAngle)) ; 
	SetDBData(hEnt, "ID",  sLegStyle+sFileNo+sLeg+"ClosingLine") ;

	// Modify polyline by moving last vertex
	nn = GetVertexCount( hCurv) ;
	SetVertex( hCurv, nn, xyThighPlt) ;
	}
    else {
	xyO.x = xyProfileLast.x  ;
	xyO.y = xyOtemplate.y ;
	hEnt = AddEntity ("line" , xyProfileLast, xyO);
   	SetDBData(hEnt, "ID", sLegStyle+sFileNo+sLeg+"ClosingLine") ;
	}

// Closing line to Template start and start
   AddEntity ("line" , xyO, xyOtemplate);
   hEnt = AddEntity ("line" , xyProfileStart, xyOtemplate);
   SetDBData(hEnt, "ID", sLegStyle+sFileNo+sLeg+"DistalClosingLine") ;
  
// Seam TRAM Lines
   PRSetLayerNotes() ;
   AddEntity("line",xyOtemplate.x, xyO.y + nSeam + 0.5, xyO.x, xyO.y + nSeam + 0.5);
   AddEntity("line", xyOtemplate.x, xyO.y + nSeam, xyO.x, xyO.y + nSeam );

// Notes about elastic 	
   SetData("TextAngle", 90);
   if (nElastic == 1)
   	AddEntity ("text", "ELASTIC", xyProfileStart.x + 0.25, xyOtemplate.y +  (xyProfileStart.y - xyOtemplate.y)/2 ) ;
   else
   	AddEntity ("text", "NO ELASTIC", xyProfileStart.x + 0.25, xyOtemplate.y +  (xyProfileStart.y - xyOtemplate.y)/2 ) ;

   if (nLegStyle == 3) // Knee Bands
   	AddEntity ("text", "ELASTIC", xyProfileLast.x - 0.375, xyO.y +  (xyProfileLast.y - xyO.y)/2 ) ;

   SetData("TextAngle", 0);

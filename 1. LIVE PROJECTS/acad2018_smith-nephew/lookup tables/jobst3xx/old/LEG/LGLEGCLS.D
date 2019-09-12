// File Name:	LGLEGCLS.D
// Purpose:	Draw closing lines for Knee High and Thigh High
//
// Version:	1.02
// Date:		25.Jan.94
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
// 06.Mar.95	GG/GD	Thigh Plate mods

   Execute ("menu", "SetLayer", hTemplateLayer);   
   if (nLegStyle == 1) {
	hEnt = AddEntity ("line", xyO, xyProfileLast) ;
   	SetDBData(hEnt, "ID", sLegStyle+sFileNo+sLeg+"ClosingLine") ;
	}
   else {
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
	nThighPltDeltaAngle = nThighPltDeltaAngle + 1 ;
	hEnt = AddEntity ("arc", xyTmp,
			      nThighPltRad,
		           	      nThighPltStartAngle,
		           	      nThighPltDeltaAngle);	
	SetDBData(hEnt, "ID",  sLegStyle+sFileNo+sLeg+"ClosingLine") ;	

	// Bottom closing line
	hEnt = AddEntity ("line", xyO, CalcXY("relpolar", xyTmp, nThighPltRad, nThighPltStartAngle)) ; 
	SetDBData(hEnt, "ID",  sLegStyle+sFileNo+sLeg+"ClosingLine") ;

	// Modify polyline by moving last vertex
//	nn = GetVertexCount( hCurv) ;
//	SetVertex( hCurv, nn, xyThighPlt) ;
	}

// Closing line to TOE
   if (xyOtemplate.x < xyToeSeam.x)
  	 AddEntity ("line" , xyO, xyToeSeam.x, xyOtemplate.y );
   else
  	 AddEntity ("line" , xyO, xyOtemplate);
   
// Seam TRAM Lines
   PRSetLayerNotes() ;
   AddEntity("line",xyToeSeam.x, xyO.y + nSeam + 0.5, xyO.x, xyO.y + nSeam + 0.5);
   AddEntity("line", xyToeSeam.x, xyO.y + nSeam, xyO.x, xyO.y + nSeam );
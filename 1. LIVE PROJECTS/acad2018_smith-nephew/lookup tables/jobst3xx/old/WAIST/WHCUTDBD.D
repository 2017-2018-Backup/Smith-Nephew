 // File Name:	WHCUTDBD.D
// Purpose:	Calculate Body points 
//		Draw Body and Cut Out for WH_CUT.D
//	
// Version:	1.03
// Date:		05.Feb.94
// Author:	Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
// 22.Sept.94	GG	Production evaluation modifications
//			Fax 6.Sept.94 
// 08.Nov.94	GG	Production evaluation modifications
//			On site
// 28.Feb.95	GG	Cut-Out cross over at waist eliminated.
//			Give warning message w.r.t Mid-Point

// Get Origin
   GetUser ("xy", "Give origin point for Cut Out", &xyOtemplate) ;
   xyO =  xyOtemplate;
   Display ("cursor", "wait", "Drawing");
//

   nDatum = int (nFoldHt / 1.5) * 1.5;
   nBodyLegTapePos = int (nFoldHt / 1.5) + 6 ;

// NB Explict leg style given
//
   PantyLeg = %false ;
   if ( nLegStyle == 1)PantyLeg = %true ;

// Calculate Cut Out and TOS
//

// For all subsequent calculations the measurements are worked
// from the right edge of the body template #8654
   nXscale = 1.25 / 1.5 ;
   nYscale = 0.5 ;
   
// Scale given values to Decimal inches
// Retain the Given value for later use as the other value will be FIGURED
//
   // Circumferences
   nGivenTOSCir = nTOSCir;
   nGivenWaistCir = nWaistCir;
   nGivenMidPointCir = nMidPointCir;
   nGivenLargestCir = nLargestCir ; 

// Start Calculating XY control points
// First Figure Heights and Circumferences
// Largest part of buttocks
   nLength = (nFoldHt + (nWaistHt - nFoldHt)  / 3) ;
   nLargestHt = FNRound((nLength - nDatum) * nXscale);
   nLargestCir = FNRound(nLargestCir * nLargestGivenRed ) ;

// Mid Point (Cleft)
   nMidPointHt =  (nLength + (nWaistHt - nFoldHt)  / 4) ;   // N.B. Carry through of nLength from above
   nMidPointHt =  FNRound((nMidPointHt - nDatum) * nXscale);
   xyMidPoint.x = xyO.x + nMidPointHt ;
   nMidPointCir = FNRound(nMidPointCir * nMidPointGivenRed);

// Fold of Buttocks (11.1)
   nFoldHt  =  FNRound((nFoldHt - nDatum) * nXscale);
   nThighCirOriginal = min ( nLeftThighCir, nRightThighCir) ;
   if (PantyLeg) { // 92% mark at thigh
	xyFoldPanty.x = xyO.x  + nFoldHt   ;
	xyFoldPanty.y = xyO.y + FNRound(nThighCirOriginal * nPantyThighRed2) * nYscale + nSeam ;
	}
   nThighCir = FNRound(nThighCirOriginal  * nThighGivenRed);

   xyFold.x = xyO.x  + nFoldHt   ;
   xyFold.y = xyO.y + nThighCir * nYscale + nSeam ;

//  Largest part of Buttocks (12.1)
   xyButtockArcCen.y  = xyO.y + (nThighCir / 2) * nYscale + nSeam ;
   xyButtockArcCen.x  = xyO.x + nLargestHt ;
   nLength = Calc ( "length", xyButtockArcCen, xyFold) ;

   xyLargest.y = xyButtockArcCen.y  + nLength  ;
   xyLargest.x = xyButtockArcCen.x ;  

// WaistCir
   nWaistCir = FNRound(nWaistCir  * nWaistGivenRed);

// TOSCir
   nTOSCir = FNRound(nTOSCir * nTOSGivenRed);

// CutOut  (13.1)
   OpenCrotch = %false;
   ClosedCrotch = %false ;
   if (StringCompare("Open Crotch", sCrotchStyle))
	OpenCrotch = %true ;
   else
	ClosedCrotch = %true ;
	
// Get distance to largest part of buttocks from seam line (13)
// nCrotchFrontFactor from WHBODDIA - VB programme

   nLength = xyLargest.y - xyO.y - nSeam ;

   xyCO_LargestBott.x = xyLargest.x ;
   xyCO_LargestBott.y = xyO.y + ((nLargestCir * nYscale - nLength) * nCrotchFrontFactor) + nSeam ;

   xyCO_LargestTop.x = xyLargest.x ;
   xyCO_LargestTop.y = xyLargest.y - ((nLargestCir * nYscale - nLength) * (1 - nCrotchFrontFactor)) ;

   nCO_HalfWidth = Calc("length", xyCO_LargestBott , xyCO_LargestTop) / 2;

// Check on Cut Out Diameter
   if ( FNRound( nCO_HalfWidth*2) > nCutOutDiaMaxTol) 
	Display ("message", "error",
			 "Warning, Cut Out Diameter is larger than 6 inches " +
			 "\nCalculated diameter is " + Format ( "length",nCO_HalfWidth*2 ));

// Cut Out 
//
// Top Of Support (if given) (Cut Out Front)
   if ( nTOSCir > 0) { 
	nDiff = int (nGivenTOSCir - nGivenLargestCir) ;
	if (nDiff < 0) nLength = nDiff * nBodyFrontReduceOff;
		else nLength = nDiff * nBodyFrontIncreaseOff;
		
	xyCO_TOSBott.y  = xyCO_LargestBott.y + nLength * nYscale ; // NB Sign carrys through
	xyCO_TOSBott.x  = xyO.x + FNRound((nTOSHt - nDatum) * nXscale);
	}

// Waist  (Cut Out Front)
   nDiff = int ( nGivenWaistCir - nGivenLargestCir ) ;
   if (nDiff < 0) nLength = nDiff * nBodyFrontReduceOff;
	   else nLength = nDiff * nBodyFrontIncreaseOff;
    xyCO_WaistBott.y  = xyCO_LargestBott.y + nLength * nYscale ; // NB Sign carrys through
    xyCO_WaistBott.x  = xyO.x + FNRound((nWaistHt - nDatum) * nXscale);   

// MidPoint  (Cut Out Front)
// Remember that mid point height was established earlier
   nDiff = int (nGivenMidPointCir - nGivenLargestCir) ;
   if (nDiff < 0) nLength = nDiff * nBodyFrontReduceOff;
   	else nLength = nDiff * nBodyFrontIncreaseOff;
   xyCO_MidPointBott.y  = xyCO_LargestBott.y + nLength * nYscale ; // NB Sign carrys through
   xyCO_MidPointBott.x  = xyO.x + nMidPointHt;   


// Cut Out (Back)
//
// Revise Cut Out Diameter to complete cut out
    nCO_Width = nCO_HalfWidth*2 * nCutOutConstructFac_1;    

// Waist  (Cut Out Back)
   xyCO_WaistTop.y  = xyCO_WaistBott.y  + nCO_Width ; 
   xyCO_WaistTop.x  = xyCO_WaistBott.x  ;

// Top Of Support   (Cut Out Back)
    xyCO_TOSTop.y  = xyCO_WaistTop.y ;
    if (nTOSCir > 0) 
	xyCO_TOSTop.x  = xyCO_TOSBott.x ;
    else 
	xyCO_TOSTop.x = xyCO_WaistTop.x + nEndBackBodyOff ;

// Back of body
// Get Back body offsets
// Check against Minimun value

   IgnoreMidPoint_CO= %false ;
   nOriginalCO_WaistTopY = xyCO_WaistTop.y  ;
   nOriginalCO_TOSTopY = xyCO_TOSTop.y  ;

   Loop = %true;
   while ( Loop ) {
	// Waist offset
	nWaistBackOff = FNRound (nWaistCir * nYscale) ;
	nWaistBackOff = (nWaistBackOff - ((xyCO_WaistTop.y - xyO.y - nSeam ) + (xyCO_WaistBott.y - xyO.y - nSeam))) / 2 ;

	// TOS Offset
	if ( nTOSCir > 0) {
		nTOSBackOff = FNRound (nTOSCir * nYscale) ;
		nTOSBackOff = (nTOSBackOff - ((xyCO_TOSTop.y - xyO.y - nSeam) + (xyCO_TOSBott.y - xyO.y - nSeam))) / 2 ;
		}
	else 
		nTOSBackOff = nWaistBackOff ;

	// MidPoint Offset
	nLength = xyCO_MidPointBott.x - xyCO_LargestBott.x ;
	aAngle = Calc ("angle", xyCO_LargestTop, xyCO_WaistTop) ;
	xyCO_MidPointTop.x = xyCO_MidPointBott.x ;
	xyCO_MidPointTop.y = xyCO_LargestTop.y  + (tan(aAngle) * nLength) ;
	nMidPointBackOff = FNRound (nMidPointCir * nYscale) ;
	nMidPointBackOff = (nMidPointBackOff  - ((xyCO_MidPointTop.y - xyO.y - nSeam) + (xyCO_MidPointBott.y - xyO.y - nSeam))) / 2 ;

	// Check that 3" distance is meet
	// Note use of 1/2 scale (nYscale) w.r.t.  nBodyBackCutOutMinTol
	Loop = %false;
	if (IgnoreMidPoint_CO) 
		nMinBackOff = FNRound(min( nWaistBackOff , nTOSBackOff ))  ;
	else
		nMinBackOff = FNRound(min( nWaistBackOff , nTOSBackOff, nMidPointBackOff  ))  ;
	if (nMinBackOff  < (nBodyBackCutOutMinTol * nYscale)) {
		Loop = %true;
		nDiff = FNRound( ((nBodyBackCutOutMinTol * nYscale) - nMinBackOff )* 2) ;
		if (nDiff == 0) break ;
		xyCO_WaistTop.y = xyCO_WaistTop.y - nDiff ;
		xyCO_TOSTop.y = xyCO_TOSTop.y - nDiff ;
		if (xyCO_WaistTop.y < xyCO_WaistBott.y) {
			IgnoreMidPoint_CO = %true;
			xyCO_WaistTop.y = nOriginalCO_WaistTopY;
			xyCO_TOSTop.y =  nOriginalCO_TOSTopY ;
			Display ("message", "error",
			 "Warning, Mid-Point ignored in calculating back Cut-Out!" +
			 "\nDistance between back of Cut-Out at Mid-Point and Profile may be less than 3\".");
			}
		}
    	} /* End while */

// Calculate back body XY points
// Waist
   xyWaist.x = xyCO_WaistTop.x ;
   xyWaist.y = xyCO_WaistTop.y + nWaistBackOff ;

// Top Of Support
   xyTOS.x = xyCO_TOSTop.x ;
   xyTOS.y = xyCO_TOSTop.y + nTOSBackOff ;

// MidPoint
   xyMidPoint.x = xyCO_MidPointTop.x ;
   xyMidPoint.y = xyCO_MidPointTop.y + nMidPointBackOff ;

// Center point of cutout (also used in crotch labeling)
   xyCO_CenterArrow.x = xyFold.x +  nCutOutConstructOff_1;  
   xyCO_CenterArrow.y = xyCO_LargestBott.y + nCO_HalfWidth ;

// Establish center, and angles of cutout arc for drawing purposes	
   xyCO_ArcCen.y = xyCO_CenterArrow.y ;
   nCO_ArcRadius =  nCO_HalfWidth ;
   xyCO_ArcCen.x = xyCO_CenterArrow.x + nCO_ArcRadius;  
   if (xyCO_ArcCen.x > xyCO_LargestTop.x ) {
   	nLength = Calc ("length", xyCO_CenterArrow, xyCO_LargestTop) / 2 ;
   	aAngle = Calc ("angle", xyCO_CenterArrow, xyCO_LargestTop);
   	nCO_ArcRadius = nLength / cos (aAngle);
  	xyCO_ArcCen.x = xyCO_CenterArrow.x + nCO_ArcRadius;  
   	aCO_ArcStart =  Calc("angle", xyCO_ArcCen, xyCO_LargestTop) ;
  	aCO_ArcDelta = Calc("angle", xyCO_ArcCen, xyCO_LargestBott) - aCO_ArcStart  ;
	}
   else	{
   	aCO_ArcStart =  90 ;
  	aCO_ArcDelta = 180  ;
	}

// Additions for OPEN Crotch
//
   if (OpenCrotch) {
	//Get Cut out Arc Center to Largest part of buttocks offset
	nArcCenToLargestOffset = xyCO_LargestTop.x -  xyCO_ArcCen.x ;
	if (nArcCenToLargestOffset < 0 ) nArcCenToLargestOffset = 0; //Make sure not -ve
	// For back / top of crotch establish length of top arc
	nTopArcSegment = (180 - aCO_ArcStart) *  %pi / 180 * nCO_ArcRadius ;
	aCO_ArcStartOriginal = aCO_ArcStart ; // Store for possible use with bottom arc
	if ( nTopArcSegment > nOpenBack) {	// End of Top O.C. lands on arc  
		xyCO_OpenArcTop = xyCO_ArcCen;
		nCO_OpenArcTopRadius = nCO_ArcRadius ;
		aCO_OpenArcTopStart = aCO_ArcStart;
		aAngle = 180 - (nOpenBack * 180) / (%pi * nCO_ArcRadius);
		aCO_OpenArcTopDelta = aAngle - aCO_OpenArcTopStart ; 
		// Modify start of Cut-Out arc
		aCO_ArcStart = aCO_OpenArcTopStart + aCO_OpenArcTopDelta;
		}
	else
		//Get distance from end of arc to lagest part of buttocks
		nTopArcToLargestOffset = nArcCenToLargestOffset - (nOpenBack - nTopArcSegment) ;
	// Do not go past Largest part of buttocks
	if (nTopArcToLargestOffset < 0) nTopArcToLargestOffset = 0; 

	// For front / bottom of crotch establish length of arc
	nBottArcSegment = ((aCO_ArcStartOriginal + aCO_ArcDelta) - 180 )  * %pi / 180 * nCO_ArcRadius;
	if ( nBottArcSegment > nOpenFront) {	// End of Bottom O.C. lands on arc  
		xyCO_OpenArcBott = xyCO_ArcCen;
		nCO_OpenArcBottRadius = nCO_ArcRadius ;
		aCO_OpenArcBottStart = 180 + (nOpenFront * 180) / (%pi * nCO_ArcRadius) ;
		aCO_OpenArcBottDelta =  aCO_ArcStartOriginal + aCO_ArcDelta - aCO_OpenArcBottStart ; 
		// Modify end of Cut-Out arc
		aCO_ArcDelta = aCO_OpenArcBottStart - aCO_ArcStart;
		}
	else	{
		//Get distance from end of arc to lagest part of buttocks
		nBottArcToLargestOffset = nArcCenToLargestOffset - (nOpenFront - nBottArcSegment) ;
		// Change aCO_ArcDelta if Start has moved (new bit)
		if ( aCO_ArcStartOriginal != aCO_ArcStart) 
			aCO_ArcDelta = aCO_ArcDelta - ( aCO_ArcStart - aCO_ArcStartOriginal) ;
		}
	// Do not go past Largest part of buttocks
	if (nBottArcToLargestOffset < 0) nBottArcToLargestOffset = 0; 

	// Establish revised arc for cut out including offset for open crotch
	// Ensure that End Points of arc do not pass largest part of buttocks

	xyPt1 = CalcXY("relpolar", xyCO_ArcCen, nCO_ArcRadius, aCO_ArcStart );
	if (xyPt1.x > xyCO_LargestTop.x) {
		xyPt1= xyCO_LargestTop;
		nTopArcToLargestOffset = 0; 
		}
	xyPt2 = CalcXY("relpolar", xyCO_ArcCen, nCO_ArcRadius, aCO_ArcStart + aCO_ArcDelta );
	if (xyPt2.x > xyCO_LargestBott.x) {
		xyPt2= xyCO_LargestBott;
		nBottArcToLargestOffset = 0; 
		}

	nError = FN_CirLinInt (xyPt1, xyPt1.x, xyPt1.y + 2,  xyCO_ArcCen, nCO_ArcRadius + nOpenOff);
	aCO_ArcStart = Calc("angle",  xyCO_ArcCen, xyInt);

	nError = FN_CirLinInt (xyPt2, xyPt2.x, xyPt2.y - 2,  xyCO_ArcCen, nCO_ArcRadius + nOpenOff);
	aCO_ArcDelta = Calc("angle",  xyCO_ArcCen, xyInt) - aCO_ArcStart;

	nCO_ArcRadius = nCO_ArcRadius + nOpenOff;

	}/* Endif OpenCrotch */
//
// DRAW Cutout and TOS
//
   PRSetLayerConstruct() ;
   hEnt = AddEntity("marker","xmarker",xyO , 0.1, 0.1) ;	
	SetDBData(hEnt,"ID", sFileNo+sLeg+"O") ;
   hEnt = AddEntity("marker","xmarker", xyTOS , 0.1, 0.1) ;	
	SetDBData(hEnt,"ID", sFileNo+sLeg+"TOS") ;
   hEnt = AddEntity("marker","xmarker", xyWaist	, 0.1, 0.1) ;	
	SetDBData(hEnt,"ID", sFileNo+sLeg+"Waist") ;
   hEnt = AddEntity("marker","xmarker", xyMidPoint , 0.1, 0.1) ;		
	SetDBData(hEnt,"ID", sFileNo+sLeg+"MidPoint") ;
   hEnt = AddEntity("marker","xmarker", xyLargest, 0.1, 0.1) ;		
	SetDBData(hEnt,"ID", sFileNo+sLeg+"Largest") ;
   hEnt = AddEntity("marker","xmarker", xyButtockArcCen, 0.1, 0.1) ;		
	SetDBData(hEnt,"ID", sFileNo+sLeg+"ButtockArcCen") ;
   hEnt = AddEntity("marker","xmarker", xyCO_TOSBott, 0.1, 0.1) ;	
	SetDBData(hEnt,"ID", sFileNo+sLeg+"CO_TOSBott") ;
   hEnt = AddEntity("marker","xmarker", xyCO_TOSTop, 0.1, 0.1) ;	
	SetDBData(hEnt,"ID", sFileNo+sLeg+"CO_TOSTop") ;
   hEnt = AddEntity("marker","xmarker", xyCO_WaistBott,0.1, 0.1) ;	
	SetDBData(hEnt,"ID", sFileNo+sLeg+"CO_WaistBott") ;
   hEnt = AddEntity("marker","xmarker", xyCO_WaistTop,0.1, 0.1) ;	
	SetDBData(hEnt,"ID", sFileNo+sLeg+"CO_WaistTop") ;
   hEnt = AddEntity("marker","xmarker", xyCO_MidPointBott,0.1, 0.1) ;	
	SetDBData(hEnt,"ID", sFileNo+sLeg+"CO_MidPointBott") ;
   hEnt = AddEntity("marker","xmarker", xyCO_MidPointTop,0.1, 0.1) ;	
	SetDBData(hEnt,"ID", sFileNo+sLeg+"CO_MidPointTop") ;
   hEnt = AddEntity("marker","xmarker", xyFold, 0.1, 0.1) ;		
	SetDBData(hEnt,"ID", sFileNo+sLeg+"Fold") ;
   if (PantyLeg) {
   	hEnt = AddEntity("marker","xmarker", xyFoldPanty, 0.1, 0.1) ;		
		SetDBData(hEnt,"ID", sFileNo+sLeg+"FoldPanty") ;
	}
   hEnt = AddEntity("marker","xmarker", xyCO_LargestTop, 0.1, 0.1) ;		
	SetDBData(hEnt,"ID", sFileNo+sLeg+"CO_LargestTop") ;
	SetDBData(hEnt,"Data",  MakeString("scalar", nTopArcToLargestOffset) );

   hEnt = AddEntity("marker","xmarker", xyCO_LargestBott, 0.1, 0.1) ;	
	SetDBData(hEnt,"ID", sFileNo+sLeg+"CO_LargestBott") ;
	SetDBData(hEnt,"Data",  MakeString("scalar", nBottArcToLargestOffset) );

   hEnt = AddEntity("marker","xmarker", xyCO_ArcCen , 0.1, 0.1) ;		
	SetDBData(hEnt,"ID", sFileNo+sLeg+"CO_ArcCen") ;
	SetDBData(hEnt,"Data",       MakeString("scalar", nCO_ArcRadius) + 	
			      " " + MakeString("scalar", aCO_ArcStart) + 
			      " " + MakeString("scalar", aCO_ArcDelta) ) ;

   if (!(xyCO_OpenArcBott.x == 0 && xyCO_OpenArcBott.y == 0) ) {
   	hEnt = AddEntity("marker","xmarker", xyCO_ArcCen , 0.1, 0.1) ;		
		SetDBData(hEnt,"ID", sFileNo+sLeg+"CO_OpenArcBott") ;
		SetDBData(hEnt,"Data",       MakeString("scalar", nCO_OpenArcBottRadius) + 	
				      " " + MakeString("scalar", aCO_OpenArcBottStart) + 
				      " " + MakeString("scalar", aCO_OpenArcBottDelta) ) ;
	}

   if (!(xyCO_OpenArcTop.x == 0 && xyCO_OpenArcTop.y == 0) ) {
   	hEnt = AddEntity("marker","xmarker", xyCO_ArcCen , 0.1, 0.1) ;		
		SetDBData(hEnt,"ID", sFileNo+sLeg+"CO_OpenArcTop") ;
		SetDBData(hEnt,"Data",       MakeString("scalar", nCO_OpenArcTopRadius) + 	
				      " " + MakeString("scalar", aCO_OpenArcTopStart) + 
				      " " + MakeString("scalar", aCO_OpenArcTopDelta) ) ;
	}

   hEnt = AddEntity("marker","xmarker", xyCO_CenterArrow, 0.1, 0.1) ;	
  	SetDBData(hEnt, "ID", sFileNo+sLeg+"CO_CenterArrow");
   	SetDBData(hEnt, "Data", MakeString("scalar",nOpenOff)) ;  

// Check for special case where the user has allowed a difference of more than 5%
// Between reductions add a marker 5% away from existing marker at fold.
   if ( abs( nLargestGivenRed - nThighGivenRed) > 0.05  ) {
	xyTmp.x = xyFold.x ;
   	xyTmp.y = xyO.y + FNRound(nThighCirOriginal  * (nThighGivenRed + 0.05)) * nYscale + nSeam ;
   	hEnt = AddEntity("marker","xmarker",xyTmp, 0.1, 0.1) ;	
  		SetDBData(hEnt, "ID", sFileNo+sLeg+"Fold+5%");
	}

   if (LeftLeg) PRSetLayerTemplateLeft () ;
	else PRSetLayerTemplateRight () ;	
//
// DRAW Cutout and TOS
//

   if ( OpenCrotch) {
	// Get Start point of arc 
	xyPt1 = CalcXY("relpolar", xyCO_ArcCen,  nCO_ArcRadius, aCO_ArcStart);
	// TOP Fillet radius
	nError = FN_CirLinInt ( xyPt1.x - nFilletRadius, xyCO_ArcCen.y, 
			     xyPt1.x - nFilletRadius, xyCO_ArcCen.y + nCO_ArcRadius,
			     xyCO_ArcCen, nCO_ArcRadius - nFilletRadius);
	xyFilletTop = xyInt;

	// TOP
	if (!(xyCO_OpenArcTop.x == 0 && xyCO_OpenArcTop.y == 0)  ){
		AddEntity ( "arc", xyCO_OpenArcTop, nCO_OpenArcTopRadius, aCO_OpenArcTopStart, aCO_OpenArcTopDelta) ;
		xyPt2 = CalcXY("relpolar",  xyCO_OpenArcTop, nCO_OpenArcTopRadius, aCO_OpenArcTopStart +  aCO_OpenArcTopDelta) ;
		AddEntity ("line", xyPt1.x, xyFilletTop.y, xyPt2) ;
		xyTmp = CalcXY("relpolar", xyCO_OpenArcTop, nCO_OpenArcTopRadius, aCO_OpenArcTopStart);
		if (xyTmp.x < xyCO_LargestTop.x) AddEntity ("line", xyTmp, xyCO_LargestTop);
		// Fillet
		AddEntity("arc", xyFilletTop, nFilletRadius, 0, Calc("angle", xyCO_ArcCen, xyFilletTop ));
		}
	else {
		nLength = xyCO_LargestTop.x - xyPt1.x - nTopArcToLargestOffset ;
		if ( nLength > 0){
			xyPt2 = CalcXY("relpolar", xyPt1, nLength, 0);
			if ( nLength > nFilletRadius) {
				xyFilletTop = xyPt1;	// gets aCO_ArcStart correct
				AddEntity ("arc", xyPt2.x - nFilletRadius, xyPt2.y - nFilletRadius, nFilletRadius, 0, 90);
				AddEntity ("line", xyPt1, xyPt2.x - nFilletRadius, xyPt2.y) ;
				AddEntity ("line", xyPt2.x,  xyPt2.y - nFilletRadius,  xyPt2.x, xyCO_LargestTop.y);
				}
			else	{
				nError = FN_CirLinInt ( xyPt2.x - nFilletRadius, xyCO_ArcCen.y, 
			   			     xyPt2.x - nFilletRadius, xyCO_ArcCen.y + nCO_ArcRadius,
			     			     xyCO_ArcCen, nCO_ArcRadius - nFilletRadius);
				xyFilletTop = xyInt;
				AddEntity("arc", xyFilletTop, nFilletRadius, 0, Calc("angle", xyCO_ArcCen, xyFilletTop ));
				AddEntity ("line", xyPt2.x, xyFilletTop.y,  xyPt2.x, xyCO_LargestTop.y);
				}
			}
		else 	{
			AddEntity("arc", xyFilletTop, nFilletRadius, 0, Calc("angle", xyCO_ArcCen, xyFilletTop ));
			AddEntity ("line", xyPt1.x, xyFilletTop.y,  xyCO_LargestTop);
			}

		if (nTopArcToLargestOffset > 0) 
			AddEntity ("line", 
				  CalcXY("relpolar", xyCO_LargestTop, nTopArcToLargestOffset, 180),  
			                  xyCO_LargestTop);
		}

	AddEntity("line", xyCO_LargestTop, xyCO_WaistTop) ;
	AddEntity("line", xyCO_WaistTop, xyCO_TOSTop) ;

	// BOTTOM
	// End points of arc 
	xyPt1 = CalcXY("relpolar", xyCO_ArcCen,  nCO_ArcRadius, aCO_ArcStart + aCO_ArcDelta);

	// BOTTOM Fillet radius initial value
	nError = FN_CirLinInt ( xyPt1.x - nFilletRadius, xyCO_ArcCen.y, 
			     xyPt1.x - nFilletRadius, xyCO_ArcCen.y - nCO_ArcRadius,
			     xyCO_ArcCen, nCO_ArcRadius - nFilletRadius);
	xyFilletBott = xyInt;

	if (!(xyCO_OpenArcBott.x  == 0 && xyCO_OpenArcBott.y  == 0) ){
		AddEntity ( "arc", xyCO_OpenArcBott, nCO_OpenArcBottRadius, aCO_OpenArcBottStart, aCO_OpenArcBottDelta) ;
		xyPt2 = CalcXY("relpolar",  xyCO_OpenArcBott, nCO_OpenArcBottRadius, aCO_OpenArcBottStart) ;
		AddEntity ("line", xyPt1.x, xyFilletBott.y, xyPt2) ;
		xyTmp = CalcXY("relpolar", xyCO_OpenArcBott, nCO_OpenArcBottRadius, aCO_OpenArcBottStart + aCO_OpenArcBottDelta);
		if (xyTmp.x < xyCO_LargestBott.x) AddEntity ("line", xyTmp, xyCO_LargestBott);
		// Fillet
		aAngle = Calc("angle", xyCO_ArcCen, xyFilletBott );
		AddEntity("arc", xyFilletBott, nFilletRadius, aAngle, 360 - aAngle );
		}
	else {
		nLength = xyCO_LargestBott.x - xyPt1.x - nBottArcToLargestOffset ;
		if ( nLength > 0){
			xyPt2 = CalcXY("relpolar", xyPt1, nLength, 0);
			if ( nLength > nFilletRadius) {
				xyFilletBott = xyPt1;	// gets aCO_ArcStart correct
				AddEntity ("arc", xyPt2.x - nFilletRadius, xyPt2.y + nFilletRadius, nFilletRadius, 270, 90);
				AddEntity ("line", xyPt1, xyPt2.x - nFilletRadius, xyPt2.y) ;
				AddEntity ("line", xyPt2.x,  xyPt2.y + nFilletRadius,  xyPt2.x, xyCO_LargestBott.y);
				}
			else	{
				nError = FN_CirLinInt ( xyPt2.x - nFilletRadius, xyCO_ArcCen.y, 
			   			     xyPt2.x - nFilletRadius, xyCO_ArcCen.y - nCO_ArcRadius,
			     			     xyCO_ArcCen, nCO_ArcRadius - nFilletRadius);
				xyFilletBott = xyInt;
				aAngle = Calc("angle", xyCO_ArcCen, xyFilletBott );
				AddEntity ("arc", xyFilletBott, nFilletRadius, aAngle, 360 - aAngle );
				AddEntity ("line", xyPt2.x, xyFilletBott.y,  xyPt2.x, xyCO_LargestBott.y);
				}
			}
		else 	{
			aAngle = Calc("angle", xyCO_ArcCen, xyFilletBott );
			AddEntity ("arc", xyFilletBott, nFilletRadius, aAngle, 360 - aAngle );
			AddEntity ("line", xyPt1.x, xyFilletBott.y,  xyCO_LargestBott);
			}

		if (nBottArcToLargestOffset > 0) 
			AddEntity ("line", 
				  CalcXY("relpolar", xyCO_LargestBott, nBottArcToLargestOffset, 180),  
			                  xyCO_LargestBott);
		}

	StartPoly ("openfitted") ;
	  AddVertex( xyCO_LargestBott);
	  AddVertex( xyCO_MidPointBott);
	  AddVertex( xyCO_WaistBott) ;
	  if ( nTOSCir > 0) AddVertex( xyCO_TOSBott) ;
	EndPoly() ;

	// Draw arc taking account of fillets
	aCO_ArcStart = Calc ("angle", xyCO_ArcCen, xyFilletTop);
	aCO_ArcDelta = Calc ("angle", xyCO_ArcCen, xyFilletBott) - aCO_ArcStart ;
	AddEntity ( "arc", xyCO_ArcCen, nCO_ArcRadius, aCO_ArcStart, aCO_ArcDelta) ;
	}
    else {
	// Closed Crotch
	AddEntity ( "arc", xyCO_ArcCen, nCO_ArcRadius, aCO_ArcStart, aCO_ArcDelta) ;
	// TOP
	AddEntity( "line", xyCO_LargestTop, xyCO_WaistTop) ;
	AddEntity( "line", xyCO_WaistTop, xyCO_TOSTop) ;
	// FIDDLY BITS, If center of cut out is below largest part of buttocks then put in joining lines 
	if (xyCO_ArcCen.x < xyCO_LargestTop.x) {
	       AddEntity( "line", CalcXY ("relpolar", xyCO_ArcCen, nCO_ArcRadius, 90), xyCO_LargestTop);
	       AddEntity( "line", CalcXY ("relpolar", xyCO_ArcCen, nCO_ArcRadius, 270), xyCO_LargestBott);
	}
	// BOTTOM
	StartPoly ("openfitted") ;
	   AddVertex (xyCO_LargestBott);
	   AddVertex( xyCO_MidPointBott);
	   AddVertex( xyCO_WaistBott) ;
	   if ( nTOSCir > 0) AddVertex( xyCO_TOSBott) ;
	EndPoly() ;
	}
	
    // Closing Line at waist or TOS
    AddEntity ("line", xyCO_TOSTop, xyTOS) ;
    if ( nTOSCir > 0) AddEntity ("line", xyCO_TOSBott,  xyTOS.x, xyO.y) ;
	else AddEntity ("line", xyCO_WaistBott, xyCO_WaistBott.x,  xyO.y) ;

    // Draw construction lines
    // 
    PRSetLayerConstruct () ;
    hEnt = AddEntity( "line", xyFold.x, xyLargest.y,
		xyFold.x, xyO.y) ;
    SetDBData (hEnt, "ID",  sFileNo + sLeg + "FoldLine" ) ; 
    // Store figured thigh cir. and actual waist cir
    SetDBData (hEnt, "Data",MakeString ("scalar",  nThighCir * nYscale + nSeam ) + " " + MakeString ("scalar", nGivenWaistCir));     	

    AddEntity( "line", xyLargest.x, xyLargest.y,
		xyLargest.x, xyO.y) ;
    AddEntity( "line", xyMidPoint.x, xyLargest.y,
		xyMidPoint.x, xyO.y) ;
    hEnt = AddEntity( "line", xyWaist.x, xyLargest.y,
		xyWaist.x, xyO.y) ;
    if (nTOSCir != 0 )
        hEnt = AddEntity( "line", xyTOS.x, xyLargest.y, xyTOS.x, xyO.y) ;
        
    SetDBData (hEnt, "ID",  sFileNo + sLeg + "EOSLine" ) ; 
    // Store figured thigh cir. and actual waist cir
    SetDBData (hEnt, "Data",MakeString ("scalar",  nThighCir * nYscale + nSeam ) + " " + MakeString ("scalar", nGivenWaistCir));     	
    
    AddEntity( "arc", xyButtockArcCen, Calc ( "length", xyButtockArcCen, xyFold) , 60, 70) ;
	
    // Add Body template tapes
    // From body position given 
    if (nBodyLegTapePos > 0) nn = nBodyLegTapePos - 1 ;
	else nn = nLastTape - 1 ;
    xyPt1.x = xyO.x ;
    xyPt1.y = xyO.y + nSeam + 0.5 ;
    while ( xyPt1.x < xyTOS.x ) {
	sSymbol = MakeString("long", nn) + "tape" ;
   	if ( !Symbol("find", sSymbol)) Exit(%cancel, "Can't find a symbol to insert\nCheck your installation, that JOBST.SLB exists\n");
	AddEntity("symbol", sSymbol,  xyPt1) ;
	xyPt1.x = xyPt1.x + 1.5 * nXscale ;
	nn = nn + 1;
	}

// Seam TRAM Lines 
   PRSetLayerNotes();
   AddEntity  ("line", xyOtemplate.x , xyOtemplate.y + nSeam + 0.5, xyPt1.x - (1.5 * nXscale ), xyPt1.y ) ; //*
   AddEntity ("line", xyOtemplate.x , xyOtemplate.y + nSeam, xyPt1.x  - (1.5 * nXscale ), xyPt1.y -0.5) ; //*

// Diagonal Fly Marks
   if (StringCompare ("Diagonal Fly", sCrotchStyle) && StringCompare("Left", sLeg)) {
   	AddEntity("marker", "closed arrow", xyCO_LargestTop, 0.5 , 0.125, 90 ) ;
   	AddEntity("marker", "closed arrow", xyCO_LargestBott, 0.5 , 0.125, 270 ) ;
	}
 
   if (LeftLeg) PRSetLayerTemplateLeft () ;
	else PRSetLayerTemplateRight () ;	
   if (nTOSCir > 0) AddEntity  ("line", xyOtemplate, xyTOS.x, xyOtemplate.y ) ;
 	else AddEntity  ("line", xyOtemplate, xyWaist.x, xyOtemplate.y ) ;

// Restore to layer 1

   Execute("menu", "SetLayer", Table("find", "layer", "1"));






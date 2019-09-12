// File Name:	WHLG2CUT.D
// Purpose:	Get Body Points and
//		Draw Body and Cut Out for Left Leg
//	
// Version:	1.04
// Date:		05.Feb.94
// Author:	Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
// 12.Mar.94	gg	Enable cutout on second leg
// 06.Oct.94	gg	Production evaluation modifications
//			Fax 6.Sept.94 
// 05.Dec.94	gg	For left leg only 
// N.B.
// xyO is set from WHLG2ORG.D

// For all subsequent calculations the measurements are worked
// from the right edge of the body template #8654
   nXscale = 1.25 / 1.5 ;
   nYscale = 0.5 ;

// Get marker details
//	
   OpenTopArcFound = %false;
   OpenBottArcFound = %false;
   sThing =sFileNo+sFirstLeg ;
   hChan = Open ("selection", "type = 'Marker'");
   if (hChan) {
	ResetSelection(hChan);
	while (hEnt = GetNextSelection(hChan)) {
		GetDBData(hEnt,"ID", &sTmp) ;
		GetGeometry(hEnt, &sType, &xyPt1, &xyTmp, &aAngle) ;
	          if ( xyPt1.x != 0 || xyPt1.y != 0) {
		if (StringCompare(sTmp,  sThing+"O"))
			GetGeometry(hEnt, &sTmp, &xyOFirstBody, &xyTmp, &aAngle) ;
		if (StringCompare(sTmp,  sThing+"TOS"))
			GetGeometry(hEnt, &sTmp, &xyTOS , &xyTmp, &aAngle) ;
		if (StringCompare(sTmp,  sThing+"Waist"))
			GetGeometry(hEnt, &sTmp, &xyWaist , &xyTmp, &aAngle) ;
		if (StringCompare(sTmp,  sThing+"MidPoint"))
			GetGeometry(hEnt, &sTmp, &xyMidPoint , &xyTmp, &aAngle) ;
		if (StringCompare(sTmp,  sThing+"Largest"))
			GetGeometry(hEnt, &sTmp, &xyLargest, &xyTmp, &aAngle) ;
		if (StringCompare(sTmp,  sThing+"ButtockArcCen"))
			GetGeometry(hEnt, &sTmp, &xyButtockArcCen, &xyTmp, &aAngle) ;
		if (StringCompare(sTmp,  sThing+"CO_LargestTop")){
			GetGeometry(hEnt, &sTmp, &xyCO_LargestTop, &xyTmp, &aAngle) ;
			GetDBData(hEnt, "Data", &sTmp);
			nTopArcToLargestOffset = Value("scalar", sTmp);
			}
		if (StringCompare(sTmp,  sThing+"CO_LargestBott")){
			GetGeometry(hEnt, &sTmp, &xyCO_LargestBott, &xyTmp, &aAngle) ;
			GetDBData(hEnt, "Data", &sTmp);
			nBottArcToLargestOffset = Value("scalar", sTmp);
			}
		if (StringCompare(sTmp,  sThing+"CO_TOSBott"))
			GetGeometry(hEnt, &sTmp, &xyCO_TOSBott, &xyTmp, &aAngle) ;
		if (StringCompare(sTmp,  sThing+"CO_TOSTop"))
			GetGeometry(hEnt, &sTmp, &xyCO_TOSTop, &xyTmp, &aAngle) ;
		if (StringCompare(sTmp,  sThing+"CO_WaistBott"))
			GetGeometry(hEnt, &sTmp, &xyCO_WaistBott, &xyTmp, &aAngle) ;
		if (StringCompare(sTmp,  sThing+"CO_WaistTop"))
			GetGeometry(hEnt, &sTmp, &xyCO_WaistTop, &xyTmp, &aAngle) ;
		if (StringCompare(sTmp,  sThing+"CO_MidPointBott"))
			GetGeometry(hEnt, &sTmp, &xyCO_MidPointBott, &xyTmp, &aAngle) ;
		if (StringCompare(sTmp,  sThing+"CO_MidPointTop"))
			GetGeometry(hEnt, &sTmp, &xyCO_MidPointTop, &xyTmp, &aAngle) ;
		if (StringCompare(sTmp,  sThing+"CO_ArcCen")){
			GetGeometry(hEnt, &sTmp, &xyCO_ArcCen, &xyTmp, &aAngle) ;
			GetDBData(hEnt, "Data", &sTmp);
			ScanLine(sTmp,"blank", &nCO_ArcRadius, &aCO_ArcStart, &aCO_ArcDelta );
			}
		if (StringCompare(sTmp,  sThing+"CO_OpenArcTop")){
			OpenTopArcFound = %true;
			GetGeometry(hEnt, &sTmp, &xyCO_OpenArcTop, &xyTmp, &aAngle) ;
			GetDBData(hEnt, "Data", &sTmp);
			ScanLine(sTmp,"blank", &nCO_OpenArcTopRadius, &aCO_OpenArcTopStart,&aCO_OpenArcTopDelta );
			}
		if (StringCompare(sTmp,  sThing+"CO_OpenArcBott")){
			OpenBottArcFound = %true;
			GetGeometry(hEnt, &sTmp, &xyCO_OpenArcBott, &xyTmp, &aAngle) ;
			GetDBData(hEnt, "Data", &sTmp);
			ScanLine(sTmp,"blank", &nCO_OpenArcBottRadius, &aCO_OpenArcBottStart,&aCO_OpenArcBottDelta );
			}
		if (StringCompare(sTmp,  sThing+"CO_CenterArrow")){
			GetGeometry(hEnt, &sTmp, &xyCO_CenterArrow, &xyTmp, &aAngle) ;
			GetDBData(hEnt, "Data", &sTmp);
			nOpenOff = Value("scalar", sTmp);
			}
		if (StringCompare(sTmp,  sThing+"Fold"))
			GetGeometry(hEnt, &sTmp, &xyFold, &xyTmp, &aAngle) ;
		if (StringCompare(sTmp,  sThing+"FoldPanty"))
			GetGeometry(hEnt, &sTmp, &xyFoldPanty, &xyTmp, &aAngle) ;
			}
		}
	}
   Close("selection", hChan); 
// Translate
// Points (Using relpolar, which is less cumbersome than TransXY)
   aTransAngle = Calc( "angle", xyOFirstBody, xyO) ;
   nTrans = Calc("length",  xyOFirstBody, xyO) ;
   aAngle = aTransAngle;

// Body Points
   xyFold = CalcXY ("relpolar", xyFold, nTrans, aAngle) ;
   xyFoldPanty = CalcXY ("relpolar", xyFoldPanty, nTrans, aAngle) ;
   xyTOS = CalcXY ("relpolar", xyTOS, nTrans, aAngle) ;
   xyWaist = CalcXY ("relpolar", xyWaist, nTrans, aAngle) ;
   xyMidPoint = CalcXY ("relpolar", xyMidPoint, nTrans, aAngle) ;
   xyLargestFirstBody =  xyLargest ;
   xyLargest = CalcXY ("relpolar", xyLargest, nTrans, aAngle) ;
   xyButtockArcCen = CalcXY ("relpolar", xyButtockArcCen, nTrans, aAngle) ;

// Translate CutOut Points
//
   xyCO_OpenArcBott = CalcXY ("relpolar", xyCO_OpenArcBott, nTrans, aAngle);
   xyCO_LargestTop = CalcXY ("relpolar", xyCO_LargestTop, nTrans, aAngle) ;
   xyCO_LargestBott = CalcXY ("relpolar", xyCO_LargestBott, nTrans, aAngle) ;
   xyCO_TOSBott = CalcXY ("relpolar", xyCO_TOSBott, nTrans, aAngle) ;
   xyCO_TOSTop = CalcXY ("relpolar", xyCO_TOSTop, nTrans, aAngle) ;
   xyCO_WaistBott = CalcXY ("relpolar", xyCO_WaistBott, nTrans, aAngle) ;
   xyCO_WaistTop = CalcXY ("relpolar", xyCO_WaistTop, nTrans, aAngle) ;
   xyCO_MidPointBott = CalcXY ("relpolar", xyCO_MidPointBott, nTrans, aAngle) ;
   xyCO_MidPointTop = CalcXY ("relpolar", xyCO_MidPointTop, nTrans, aAngle) ;
   xyCO_ArcCen = CalcXY ("relpolar", xyCO_ArcCen, nTrans, aAngle) ;
   xyCO_CenterArrow = CalcXY ("relpolar", xyCO_CenterArrow, nTrans, aAngle) ;
   xyCO_OpenArcTop = CalcXY ("relpolar", xyCO_OpenArcTop, nTrans, aAngle) ;

//
// DRAW Cutout
//
// Always Left Leg 
   PRSetLayerTemplateLeft () ;	
   hTemplateLayer = Table("find", "layer", "TemplateLeft");

   if (OpenCrotch) {
// Get Start point of arc 
	xyPt1 = CalcXY("relpolar", xyCO_ArcCen,  nCO_ArcRadius, aCO_ArcStart);
	// TOP Fillet radius
	nError = FN_CirLinInt ( xyPt1.x - nFilletRadius, xyCO_ArcCen.y, 
			     xyPt1.x - nFilletRadius, xyCO_ArcCen.y + nCO_ArcRadius,
			     xyCO_ArcCen, nCO_ArcRadius - nFilletRadius);
	xyFilletTop = xyInt;

	// TOP
	if (OpenTopArcFound){
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

	if (OpenBottArcFound ){
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
	
    // Closing Line at waist or TOS & fold lines
    AddEntity ("line", xyCO_TOSTop, xyTOS) ;
    if ( nTOSCir > 0) {
	AddEntity ("line", xyCO_TOSBott,  xyTOS.x, xyO.y) ;
	AddEntity  ("line", xyO, xyTOS.x, xyO.y ) ;
	}
    else{
	AddEntity  ("line", xyO, xyWaist.x, xyO.y ) ;
	AddEntity ("line", xyCO_WaistBott, xyCO_WaistBott.x,  xyO.y) ;
	}

    // Draw construction lines
    // 
    PRSetLayerConstruct () ;
    hEnt = AddEntity( "line", xyFold.x, xyLargest.y,
		xyFold.x, xyO.y) ;
    SetDBData (hEnt, "ID",  sFileNo + sLeg + "FoldLine" ) ; 
    // Store figured thigh cir. and actual waist cir
    SetDBData (hEnt, "Data",MakeString ("scalar",  nThighCir) + " " + MakeString ("scalar", nGivenWaistCir));     	

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
    SetDBData (hEnt, "Data",MakeString ("scalar",  nThighCir) + " " + MakeString ("scalar", nGivenWaistCir));     	

    AddEntity( "arc", xyButtockArcCen, Calc ( "length", xyButtockArcCen, xyFold) , 60, 70) ;
    
    // markers
    hEnt = AddEntity("marker","xmarker", xyCO_WaistBott,0.1, 0.1) ;	
	SetDBData(hEnt,"ID", sFileNo+sLeg+"CO_WaistBott") ;

// Add Body template tapes
// From body position given
//
    if (nBodyLegTapePos > 0) nn = nBodyLegTapePos - 1 ;
	else nn = nLastTape - 1 ;

    xyPt1.x = xyO.x ;
    xyPt1.y = xyO.y + nSeam + 0.5 ;
    while ( xyPt1.x < xyTOS.x ) {
	sSymbol = MakeString("long", nn) + "tape" ;
   	if ( !Symbol("find", sSymbol)) Exit(%cancel, "Can't find a symbol to insert\nCheck your installation, that JOBST.SLB exists");
	AddEntity("symbol", sSymbol,  xyPt1) ;
	xyPt1.x = xyPt1.x + 1.5 * nXscale ;
	nn = nn + 1;
	}
// Draw Seam Tram Lines
//
   PRSetLayerNotes();
   AddEntity ("line", xyO.x , xyO.y + nSeam, xyPt1.x - (1.5 * nXscale ), xyPt1.y -0.5) ;
   AddEntity  ("line", xyO.x , xyO.y + nSeam + 0.5, xyPt1.x - (1.5 * nXscale ), xyPt1.y ) ;

// Diagonal Fly Marks
   if (StringCompare ("Diagonal Fly", sCrotchStyle)) {
   	AddEntity("marker", "closed arrow", xyCO_LargestTop, 0.5 , 0.125, 90 ) ;
   	AddEntity("marker", "closed arrow", xyCO_LargestBott, 0.5 , 0.125, 270 ) ;
	}

// Crotch arrow
   hEnt = AddEntity("marker","xmarker", xyCO_CenterArrow, 0.1, 0.1) ;	
   SetDBData(hEnt, "ID", sFileNo+sLeg+"CO_CenterArrow");
   SetDBData(hEnt, "Data", MakeString("scalar",nOpenOff)) ;  





// File Name:	WHLG2DBD.D
// Purpose:	Get Body Points
//		Draw Body for Right leg only
//	
// Version:	1.01
// Date:		05.Dec.94
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action

// N.B.
// xyO is set from WHLG2ORG.D

// For all subsequent calculations the measurements are worked
// from the right edge of the body template #8654
   nXscale = 1.25 / 1.5 ;
   nYscale = 0.5 ;

// Get marker details
//	
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

//
// DRAW Body  closing lines and Constuction lines 
//
// Always Right leg
   PRSetLayerTemplateRight () ;	
   hTemplateLayer = Table("find", "layer", "TemplateRight");

    // Closing Line at waist or TOS & Foldlines
    AddEntity ("line", xyTOS,  xyTOS.x, xyO.y) ;
    AddEntity  ("line", xyO, xyTOS.x, xyO.y ) ;
    if ( nTOSCir == 0) {
    	PRSetLayerNotes();
    	AddEntity ("line", xyWaist, xyWaist.x, xyO.y);
	}

    // Draw construction lines
    // 
    PRSetLayerConstruct () ;
    hEnt = AddEntity( "line", xyFold.x, xyLargest.y,
		xyFold.x, xyO.y) ;
    SetDBData (hEnt, "ID",  sFileNo + sLeg + "FoldLine" ) ; 
    // Store figured thigh cir. and actual waist cir
    SetDBData (hEnt, "Data",MakeString ("scalar",  nThighCir ) + " " + MakeString ("scalar", nGivenWaistCir));     	

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





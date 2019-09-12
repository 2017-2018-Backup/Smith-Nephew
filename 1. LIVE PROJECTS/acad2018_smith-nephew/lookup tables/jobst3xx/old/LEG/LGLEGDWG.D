// File Name:	LGLEGDWG.D
// Purpose:	Draw leg and foot
//
// Version:	1.06
// Date:		25.Jan.94
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
// 12.Mar.94	GG	Modifications wrt JOBST comments
//			Letter dated 4/3/94 parts 1,5
//			Fax dated 18/2/94  parts 2,3,7
// 20.Jul.94	GG	G.Dunne Visit modifications
// 28.Sep.94	GG	Production Evaluation Modifications
//			Fax 9.Sept.94
//			//* - Small Heel mods
//			//** - Big Heel mods
// 08.Nov.94	GG	Onsite Production evaluation
// 16.Nov.94	GG	Enable pleats greater than 1.25"
// 04.Dec.94	GG	JOBSTEX Gradient Fabric Class
// 09.Dec.94	GG	Modified to draw for Anklet, Knee and Thigh High
// 06.Mar.95	GG/GD	Thigh Plate mods.

//			

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

// Draw for Straight & Straight types
   if (    StringCompare("Straight", sToeStyle)
        || StringCompare ("Soft Enclosed B/M", sToeStyle) ) {
	// Draw for straight & Straight types
	StartPoly ("fitted") ;
	AddVertex (xyToeOFF) ;
   	}
   else {
	// Draw Toe Curve
	StartPoly ("fitted") ;
	AddVertex (xyToeSeam) ;
	// First toe curve
	aAngle = Calc("angle", xyToeCntrLow ,xyToeCntrMid) ;
	aPrevAngle =  Calc("angle", xyToeCntrLow ,xyToeSeam) ;
	if (aAngle > aPrevAngle ) aAngleInc = (aAngle - aPrevAngle) / 3 ;
		else aAngleInc = ((aAngle + 360) - aPrevAngle) / 3 ;
	ii = 1 ;
	while ( ii  <= 3) {
		AddVertex (CalcXY ("relpolar", xyToeCntrLow , nToeLowR, aPrevAngle + aAngleInc * ii));
		ii = ii + 1 ;
		}
	// Toe Point curve
	aAngle = Calc("angle", xyToeCntrMid ,xyToeCntrLow) ;
	aPrevAngle =  Calc("angle", xyToeCntrHigh ,xyToeCntrMid) ;
	if (aAngle > aPrevAngle ) aAngleInc = (aAngle - aPrevAngle) / 3 ;
		else aAngleInc = ((aAngle + 360) - aPrevAngle) / 3 ;
	ii = 1 ;
	xyToePnt = CalcXY ("relpolar", xyToeCntrMid , nToeMidR, aAngle - aAngleInc * ii);
	while ( ii  <= 3) {
		AddVertex (CalcXY ("relpolar", xyToeCntrMid , nToeMidR, aAngle - aAngleInc * ii));
		ii = ii + 1 ;
		}

	// Top of toe curve
	aAngle = Calc("angle", xyToeCntrHigh ,xyToeCntrMid) ;
	aPrevAngle =  Calc("angle", xyToeCntrHigh ,xyAnkleMPrev) ;
	if (aPrevAngle <= 90) aPrevAngle = 90;
	if (aAngle > aPrevAngle ) aAngleInc = (aAngle - aPrevAngle) / 3 ;
		else aAngleInc = ((aAngle + 360) - aPrevAngle) / 3 ;
	ii = 1 ;
	while ( ii  <= 2) {
		AddVertex (CalcXY ("relpolar", xyToeCntrHigh , nToeHighR, aAngle - aAngleInc * ii));
		ii = ii + 1 ;
		}
	if (!SmallHeel) AddVertex (xyAnkleMPrev) ;
	}

// Draw Heel
// Add Start of heel
   if ( xyAnkleMPrev.y != 0 && !SmallHeel ) AddVertex ( xyAnkleM)  ;	//*

   aPrevAngle = 270 ;
   aAngle =  Calc("angle", xyHeelCntrDistal ,xyHeelCntrMidDistal) ;
   if (aAngle > aPrevAngle ) aAngleInc = (aAngle - aPrevAngle) / 3 ;
	else aAngleInc = ((aAngle + 360) - aPrevAngle) / 3 ;
   ii = 1 ;
   while ( ii  <= 2) {
	AddVertex (CalcXY ("relpolar", xyHeelCntrDistal , nHeelR1, aPrevAngle + aAngleInc * ii));
	ii = ii + 1 ;
	}
	
   aPrevAngle = 90 ;
   aAngle =  Calc("angle" ,xyHeelCntrMidDistal , xyHeelCntrDistal) ;
   if (aAngle > aPrevAngle ) aAngleInc = (aAngle - aPrevAngle) / 3 ;
	else aAngleInc = ((aAngle + 360) - aPrevAngle) / 3 ;
   ii = 1 ;
   while ( ii  <= 3) {
	AddVertex (CalcXY ("relpolar", xyHeelCntrMidDistal , nHeelR2, aAngle - aAngleInc * ii));
	ii = ii + 1 ;
	}

   if (!SmallHeel || BigAnkle) AddVertex (xyHeel) ;	//*

   aAngle = 90 ;
   aPrevAngle =  Calc("angle" ,xyHeelCntrMidProximal , xyHeelCntrProximal) ;
   if (aAngle > aPrevAngle ) aAngleInc = (aAngle - aPrevAngle) / 3 ;
	else aAngleInc = ((aAngle + 360) - aPrevAngle) / 3 ;
   if (BigAnkle) nPts = 0; //**
	else nPts = 2; //**
   ii = 0 ;
   while ( ii  <= nPts) {
	AddVertex ( CalcXY ("relpolar", xyHeelCntrMidProximal , nHeelR2, aAngle - aAngleInc * ii)) ; 
	ii = ii + 1 ;
	}
	
   aAngle = 270 ;
   aPrevAngle =  Calc("angle" , xyHeelCntrProximal, xyHeelCntrMidProximal ) ;
   if (aAngle > aPrevAngle ) aAngleInc = (aAngle - aPrevAngle) / 3 ;
	else aAngleInc = ((aAngle + 360) - aPrevAngle) / 3 ;
   if (BigAnkle) nPts = 0 ; //**
	else nPts = 2; //**
   ii = 1 ;
  while ( ii  <= nPts) {	//*
	xyTmp = CalcXY ("relpolar", xyHeelCntrProximal , nHeelR3, aPrevAngle + aAngleInc * ii); //*
	if (xyTmp.x < xyAnkle.x) AddVertex ( xyTmp); //*
	ii = ii + 1 ;
	}
   // End of drawing of Foot	
//
// Draw Leg 
//
   PROpenTemplateFile ();

// Skip to FirstTape
    nn = 1;
    while (nn < nFirstTape  ) {
	GetLine(hChan, &sLine) ;
	nn = nn + 1;
	} 

// Skip to ankletape 
    xyTmp = xyOtemplate ;
    while ( nn < nAnkleTape ) {
	GetLine(hChan, &sLine) ;
	ScanLine( sLine, "blank", &nNo, &sScale, &nSpace, &n20Len, &nReduction);
	if ((nn == nStyleFirstTape) && (nFootPleat1 !=0)) nSpace =  nFootPleat1 ;
	if ((nn == nStyleFirstTape+1) && (nFootPleat2 !=0)) nSpace = nFootPleat2 ;
    	xyTmp.x = xyTmp.x + nSpace ;
	nn = nn + 1;
	} 

    nn = nAnkleTape  ;

    while (nn <= nStyleLastTape ) {
	nTapeLen = FNRound(FNDecimalise(FNGetTape(nn)) * nUnitsFac) ;
	if ( nFabricClass == 2) nLength = (nTapeLen * (100 -  Value("scalar",StringMiddle(sReduction,( (nn - 1)*3)+1,3))) / 100) /  2 ;
		else nLength = n20Len / 20 * nTapeLen ;
	nRedStep = n20Len /(20 * 8) ;		

	if ( nLegStyle == 0  && nn == nStyleLastTape)
		// For Anklets release last tape to a 14 reduction
		// by counting out
		nLength = nLength + ((nReduction - 14 )  * nRedStep) ;
		
	if ( nLegStyle == 1 && nn == nStyleLastTape)
	                // For Knee length
		// Release last tape to a given reduction
		nLength = (nTapeLen * ((100 - nStyleLastRed ) /100) ) / 2 ;
		
	if ( nLegStyle == 2  && nn == nStyleLastTape)
		// For Thigh bands release last tape to a given reduction
		// by counting out
		nLength = nLength + ((nReduction - nStyleLastRed )  * nRedStep) ;

	if (nn == nAnkleTape && nFabricClass != 2 )
		// Release the ANKLE tape to the CALCULATED reduction
		nLength = nLength + ((nReduction - nReductionAnkle) * nRedStep) ;

	xyTmp.y = xyOtemplate.y + nSeam + nLength ;
	AddVertex (xyTmp) ; 	
	if (nFabricClass == 2) {
   		SetData("TextHorzJust", 2);		// Center
 		sSymbol = MakeString("long", nNo) + "tape" ;
		if ( !Symbol("find", sSymbol)) Exit(%cancel, "Can't find a symbol to insert\nCheck your installation, that JOBST.SLB exists");
		AddEntity("symbol", sSymbol,  xyTmp) ;
		AddEntity ("text", Format("length", nTapeLen), xyTmp.x, xyTmp.y - 0.5) ;
		AddEntity ("text", StringMiddle(sReduction,( (nn - 1)*3)+1,3), xyTmp.x, xyTmp.y - 0.7) ;
		AddEntity ("text", StringMiddle(sStretch,( (nn - 1)*3)+1,3), xyTmp.x , xyTmp.y - 0.9) ;
		AddEntity ("text", StringMiddle(sTapeMMs,( (nn- 1)*3)+1,3),xyTmp.x, xyTmp.y - 1.1) ;
		}

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

    hChan = Open ("selection", "layer = 'Construct' AND type = 'Curve'");
    if( hChan ) {
	ResetSelection (hChan);
	hCurv = GetNextSelection(hChan);
	SetEntityData ( hCurv, "layer", hTemplateLayer);
   	SetDBData(hCurv, "ID",sLegStyle+sFileNo+sLeg+"LegCurve") ;
	}
    Close ("selection", hChan);

// Draw foot points
   PRSetLayerConstruct();
   hEnt =  AddEntity ("marker", "xmarker", xyAnkle, 0.2, 0.2) ;
	SetDBData(hEnt, "ID", sLegStyle+sFileNo+sLeg+"Ankle");
	sTmp = MakeString("scalar",  xyHeelCntrProximal.x - xyAnkle.x ) + " " +  MakeString("scalar", xyHeelCntrProximal.y - xyAnkle.y  )  ;
	SetDBData(hEnt, "Data", sTmp);
   hEnt = AddEntity ("marker", "xmarker", xyHeel, 0.2, 0.2) ;
	SetDBData(hEnt, "ID", sLegStyle+sFileNo+sLeg+"Heel");
	SetDBData(hEnt, "Data", MakeString("long", SmallHeel));
   hEnt = AddEntity ("marker", "xmarker", xyAnkleM, 0.2, 0.2) ;
	SetDBData(hEnt, "ID", sLegStyle+sFileNo+sLeg+"AnkleM");	
	sTmp = MakeString("scalar",  xyHeelCntrDistal.x - xyAnkleM.x ) + " " +  MakeString("scalar", xyHeelCntrDistal.y - xyAnkleM.y  )  ;
	SetDBData(hEnt, "Data", sTmp);
   if ( nAge > 10) AddEntity ("marker", "xmarker", xyAnkleMPrev, 0.2, 0.2) ;
   if (SmallHeel) AddEntity ("arc",   xyHeelCntrProximal, nHeelR3, 180, 90) ; //**


// Draw rest of it
   Execute ("menu", "SetLayer", hTemplateLayer);

// Add Closing lines at TOE
   if (xyOtemplate.x > xyToeSeam.x)
   	AddEntity ("line" , xyOtemplate, xyToeSeam.x, xyOtemplate.y); 
   AddEntity ("line" , xyToeSeam.x, xyOtemplate.y, xyToeSeam);

// Toe endings
    xyToeCL.x = xyToeSeam.x ;
    xyToeCL.y = xyOtemplate.y ;

    if ( StringCompare ("Soft Enclosed", sToeStyle) && !nFootLength && (nAge <= 10)) {
	AddEntity ("line", xyToeSeam, xyToeSeam.x, xyAnkleM.y);
	}

    if ( StringCompare ("Cut-Back", sToeStyle, 8)
         ||  ( StringCompare ("Soft Enclosed", sToeStyle) && !nFootLength && (nAge > 10)) ) {
	nRightOffset = 0.75 ;
	nLeftOffset = 0.25 ;
	xyTmp.x = xyToeCntrMid.x - nToeMidR ;
	xyTmp.y =  xyToeCntrMid.y ;
	aAngle = Calc("angle", xyToeCntrLow, xyTmp ) ;
	aPrevAngle = Calc("angle", xyToeCntrLow, xyToeCL) ;
   	if (aAngle > aPrevAngle ) aAngleInc = (aAngle - aPrevAngle) / 2 ;
		else aAngleInc = ((aAngle + 360) - aPrevAngle) / 2 ;
	AddEntity("poly", 
		"openfitted",
		 xyTmp,
		CalcXY( "relpolar", xyToeCntrLow, nToeLowR - nLeftOffset, aPrevAngle + aAngleInc),
		xyToeCL) ;
	PRSetLayerNotes();
	AddEntity("poly", 
		"openfitted",
		 xyTmp,
		CalcXY( "relpolar", xyToeCntrLow, nToeLowR + nRightOffset, aPrevAngle + aAngleInc),
		xyToeCL) ;
	}

    if ( StringCompare ("Soft Enclosed", sToeStyle) && nFootLength) {
 	if (nAge > 10) nRightOffset = 2.25 ;
		else nRightOffset = 1.75 ;
	sFootLabel = Format("length", nRightOffset) + " Soft Enclosed" ;
	xyTmp.x = xyToeCntrMid.x - nToeMidR ;
	xyTmp.y =  xyToeCntrMid.y ;
	aAngle = Calc("angle", xyToeCntrLow, xyTmp ) ;
	aPrevAngle = Calc("angle", xyToeCntrLow, xyToeCL) ;
	if (aAngle > aPrevAngle ) aAngleInc = (aAngle - aPrevAngle) / 2 ;
		else aAngleInc = ((aAngle + 360) - aPrevAngle) / 2 ;
   	Execute ("menu", "SetLayer", hTemplateLayer);
	xyTmp = CalcXY( "relpolar", xyToeCntrLow, nToeLowR, aPrevAngle + aAngleInc) ;
	AddEntity ("line",  xyTmp.x + nRightOffset,  
		             xyTmp.y,
		             xyTmp.x + nRightOffset, 
		             xyAnkleM.y) ;
	AddEntity ("line",  xyTmp.x + nRightOffset,  
		             xyTmp.y,
		             xyTmp) ;
	}

   if (   StringCompare("Straight", sToeStyle) || StringCompare ("Soft Enclosed B/M", sToeStyle) ) {
	AddEntity("Line", xyToeSeam, xyToeOFF) ;
	}
   else {
	PRSetLayerNotes();
	AddEntity ("line", xyToeSeam.x, xyFirstTape.y, xyToePnt.x, xyFirstTape.y) ;
	}

// Foot lable
   PRSetLayerNotes() ;
   AddEntity ("text", sFootLabel, xyToeSeam.x + 1.5, xyToeSeam.y + 1) ;


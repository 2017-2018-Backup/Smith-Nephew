// File Name:	WHLGCDWG.D
// Purpose:	Draw leg and foot (CHAP STYLE)
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
// 28.Sept.94	GG	Production Evaluation Modifications
//			Fax 9.Sept.94
//			//* - Small Heel mods
//			//** - Big Heel mods
// 08.Nov.94	GG	Onsite Production evaluation
// 16.Nov.94	GG	Enable pleats greater than 1.25"
// 04.Dec.94	GG	JOBSTEX Gradient Fabric Class
// 30.Sep.96	GG	Modified for use in drawing chap style

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

// Check if a panty leg if not draw foot

if (!FootLess) {
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
   }
 else
   { 
   // Draw as a panty leg
   StartPoly ("fitted") ;
   if ( (nLastTape -  nFirstTape) < 3 ) ShortLeg = %true;
   }
//
// Draw Leg and Back Body Profile
//
   PROpenTemplateFile ();

// Skip to FirstTape
    nn = 1;
    while (nn < nFirstTape  ) {
	GetLine(hChan, &sLine) ;
	nn = nn + 1;
	} 

// Skip to ankle tape 
    xyTmp = xyOtemplate ;
    while ( nn < nAnkleTape ) {
	GetLine(hChan, &sLine) ;
	ScanLine( sLine, "blank", &nNo, &sScale, &nSpace, &n20Len, &nReduction);
	if ((nn == nFirstTape) && (nFootPleat1 !=0)) nSpace =  nFootPleat1 ;
	if ((nn == nFirstTape+1) && (nFootPleat2 !=0)) nSpace = nFootPleat2 ;
    	xyTmp.x = xyTmp.x + nSpace ;
	nn = nn + 1;
	} 

     if (FootLess) {
	nn = nFirstTape;
	ScanLine( sLine, "blank", &nNo, &sScale, &nSpace, &n20Len, &nReduction);
	}
     else
	nn = nAnkleTape  ;

    while (nn <= nLastTape ) {
	nTapeLen = FNRound(FNDecimalise(FNGetTape(nn)) * nUnitsFac) ;
	if ( nFabricClass == 2) nLength = (nTapeLen * (100 -  Value("scalar",StringMiddle(sReduction,( (nn - 1)*3)+1,3))) / 100) /  2 ;
		else nLength = n20Len / 20 * nTapeLen ;
	nRedStep = n20Len /(20 * 8) ;
	if ((nn == nFirstTape) && FootLess ) {
		// Release the distal tape  to a 95% reduction,  92% reduction if +3 or +1-1/2 tape
		if ( nFirstTape > 8) nLength = (nTapeLen * 0.95) / 2 ;
		        else nLength = (nTapeLen * 0.92) / 2 ;
		xyLastTape.x = xyOtemplate.x ;
		xyLastTape.y = xyOtemplate.y + nSeam + nLength ;
		}
	if (nn == nAnkleTape && nFabricClass != 2 )
		// Release the ANKLE tape to the CALCULATED reduction
		nLength = nLength + ((nReduction - nReductionAnkle) * nRedStep) ;
		
	if (nn == nLastTape && nFabricClass != 2 )
		// Release last tape to a 14 reduction
		nLength = nLength + ((nReduction - 14 )  * nRedStep) ;

	xyTmp.y = xyOtemplate.y + nSeam + nLength ;
	if (!ShortLeg) AddVertex (xyTmp) ; 

	if (ShortLeg && nn == nFirstTape) xyPt1 = xyTmp;
	if (ShortLeg && nn == nFirstTape+1) xyPt2 = xyTmp;
	if (ShortLeg && nn == nFirstTape+2) xyPt3 = xyTmp;

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

	nn = nn + 1;
	GetLine(hChan, &sLine) ;
	ScanLine( sLine, "blank", &nNo, &sScale, &nSpace, &n20Len, &nReduction);
	if ((nn == nFirstTape+1) && (nFootPleat1 !=0)) nSpace = nFootPleat1 ;
	if ((nn == nFirstTape+2) && (nFootPleat2 !=0)) nSpace = nFootPleat2 ;
	if ((nn == nLastTape) && (nTopLegPleat1 !=0)) nSpace =  nTopLegPleat1 ;
	if ((nn == nLastTape-1) && (nTopLegPleat2 !=0)) nSpace =  nTopLegPleat2 ;
	xyTmp.x = xyTmp.x + nSpace ;
	}
    Close ("file", hChan) ;

   EndPoly () ;
    // Shortleg Fudge
    //
    if (ShortLeg) 
	AddEntity("poly", 
	"openfitted",
	 xyPt1,
	 xyPt2,
	 xyPt3) ;
	

    // Get polyline entity handle
    // Change layer and set DB values
     
     hCurv = UID("find", UID("getmax")) ;	
     SetEntityData ( hCurv, "layer", hTemplateLayer);
     SetDBData(hCurv, "ID", sID+"LegCurve") ;

	
if (!FootLess) {
// Draw foot points
   PRSetLayerConstruct();
   hEnt =  AddEntity ("marker", "xmarker", xyAnkle, 0.2, 0.2) ;
	SetDBData(hEnt, "ID", sID+"Ankle");
	sTmp = MakeString("scalar",  xyHeelCntrProximal.x - xyAnkle.x ) + " " +  MakeString("scalar", xyHeelCntrProximal.y - xyAnkle.y  )  ;
	SetDBData(hEnt, "Data", sTmp);
   hEnt = AddEntity ("marker", "xmarker", xyHeel, 0.2, 0.2) ;
	SetDBData(hEnt, "ID", sID+"Heel");
	SetDBData(hEnt, "Data", MakeString("long", SmallHeel));
   hEnt = AddEntity ("marker", "xmarker", xyAnkleM, 0.2, 0.2) ;
	SetDBData(hEnt, "ID", sID+"AnkleM");	
	sTmp = MakeString("scalar",  xyHeelCntrDistal.x - xyAnkleM.x ) + " " +  MakeString("scalar", xyHeelCntrDistal.y - xyAnkleM.y  )  ;
	SetDBData(hEnt, "Data", sTmp);
   if ( nAge > 10) AddEntity ("marker", "xmarker", xyAnkleMPrev, 0.2, 0.2) ;
   if (SmallHeel) AddEntity ("arc",   xyHeelCntrProximal, nHeelR3, 180, 90) ; //**


// Draw rest of it
   Execute ("menu", "SetLayer", hTemplateLayer);

// Add Closing lines at TOE
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
// Seam TRAM Lines
   AddEntity("line",xyToeSeam.x, xyOtemplate.y + nSeam + 0.5,  xyOtemplate.x,  xyOtemplate.y + nSeam + 0.5);
   AddEntity("line", xyToeSeam.x,  xyOtemplate.y + nSeam,  xyOtemplate.x,  xyOtemplate.y + nSeam );
   }
else {
   // Draw seam line	and panty line
   Execute ("menu", "SetLayer", hTemplateLayer);
   hEnt = AddEntity ("line", xyLastTape, xyOtemplate);
   SetDBData(hEnt, "ID", sID+"DistalClosingLine") ;
   PRSetLayerNotes() ;
   if (nElastic == 1 ) sFootLabel = "ELASTIC" ;
   if (nElastic == -1) sFootLabel = "NO ELASTIC" ;
   // Add the text as a symbol (For latter use by the StumpTool)
   if ( !Symbol("find",  "TextAsSymbol")) Display("message","error", "Elastic at distal end of support not added.\nCheck and add text manually");
   hEnt = AddEntity("symbol", "TextAsSymbol", xyLastTape.x + 0.25, xyOtemplate.y +  (xyLastTape.y - xyOtemplate.y)/2, 1, 1, 90 ) ;
   SetDBData(hEnt, "ID", sID+"PantyElasticNote") ;
   SetDBData(hEnt, "Data", sFootLabel ) ;
   }


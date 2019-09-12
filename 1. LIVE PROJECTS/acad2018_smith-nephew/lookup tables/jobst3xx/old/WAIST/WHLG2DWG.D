// File Name:	WHLG2DWG.D
// Purpose:	Draw Second Leg profile
//		draws foot
//		Add Patient data
//
// Version:	1.02
// Date:		07.Feb.94
// Author:	Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
// See WHLG1DWG.D

   PRSetLayerConstruct () ;

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
	AddVertex (xyAnkleMPrev) ;
	}

// Draw Heel
// Add Start of heel
   if ( xyAnkleMPrev.y != 0 && !SmallHeel) AddVertex ( xyAnkleM)  ;

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

   if( !SmallHeel || BigAnkle) AddVertex (xyHeel) ;

   aAngle = 90 ;
   aPrevAngle =  Calc("angle" ,xyHeelCntrMidProximal , xyHeelCntrProximal) ;
   if (aAngle > aPrevAngle ) aAngleInc = (aAngle - aPrevAngle) / 3 ;
	else aAngleInc = ((aAngle + 360) - aPrevAngle) / 3 ;
   if (BigAnkle) nPts = 0;
	else nPts = 2;
   ii = 0 ;
   while ( ii  <= nPts) {
	AddVertex (CalcXY ("relpolar", xyHeelCntrMidProximal , nHeelR2, aAngle - aAngleInc * ii));
	ii = ii + 1 ;
	}
	
   aAngle = 270 ;
   aPrevAngle =  Calc("angle" , xyHeelCntrProximal, xyHeelCntrMidProximal ) ;
   if (aAngle > aPrevAngle ) aAngleInc = (aAngle - aPrevAngle) / 3 ;
	else aAngleInc = ((aAngle + 360) - aPrevAngle) / 3 ;
   if (BigAnkle) nPts = 0;
	else nPts = 2;
   ii = 1 ;
   while ( ii  <= nPts) {
	xyTmp = CalcXY ("relpolar", xyHeelCntrProximal , nHeelR3, aPrevAngle + aAngleInc * ii);
	if (xyTmp.x < xyAnkle.x) AddVertex (xyTmp);
	ii = ii + 1 ;
	}
   // End of drawing of Foot	
   }
 else
   { 
   // Draw as a panty leg
   StartPoly ("fitted") ;
   }
//
// Draw Leg and Back Body Profile
//
// Load template data file
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
	if ( xyTmp.x >= xyFold.x){
		xyPt2 = xyTmp;	// Store this for later use
		break;
		}
 	else {
		AddVertex (xyTmp) ;
		xyPt1 = xyTmp;	// Store this for later use
		}

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

    // check if we use last leg tape or the fold position
    if (nLegStyle == 1) AddVertex (xyFoldPanty);  // We must hit this point for panties
	else  if (xyTmp.y < xyFold.y) AddVertex (xyFold);
 		else  AddVertex (xyTmp) ;

     AddVertex ( xyLargest) ;

// Loop through first leg data file looking for changes that have been made to the curve
// after the largest part of buttocks on the first drawn body 
// nTrans and aTransAngle from WHLG2DBD.D
 
   hFileCurve = Open ("file", "C:\\JOBST\\LEGCURVE.DAT", "readonly") ;
   ii = 1 ;
   AddPoint = %false;
   while ( ii <= nLegVertexCount ) {
	GetLine(hFileCurve, &sLine) ;
	ScanLine(sLine,"blank", &nX, &nY) ;
	if (nX > xyLargestFirstBody.x) AddPoint = %true ;
	if (AddPoint) AddVertex (CalcXY ("relpolar", nX, nY, nTrans, aTransAngle)) ;
	ii = ii + 1;
	}
   Close ("file", hFileCurve) ;  
   EndPoly();

    // Get polyline entity handle
    // Change layer and set DB values
     hCurv = UID("find", UID("getmax")) ;	
     SetEntityData ( hCurv, "layer", hTemplateLayer);
     SetDBData(hCurv, "ID", sID+"LegCurve") ;
  
   if (nTOSCir == 0){
   	Execute ("menu", "SetLayer", hTemplateLayer);
	AddEntity ("line",  xyWaist, xyTOS) ;
	}
 

if (!FootLess) {
// Draw foot points
   PRSetLayerConstruct () ;	
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

// Add Fold point for use in editor
   if (nFabricClass == 2){
   	hEnt = AddEntity("marker","xmarker", xyFold, 0.1, 0.1) ;		
	SetDBData(hEnt,"ID", sID+"Fold") ;
	}

// Draw rest of it
   Execute ("menu", "SetLayer", hTemplateLayer);

// Add Closing lines at TOE
  AddEntity ("line" , xyO, xyToeSeam.x, xyOtemplate.y);
  AddEntity ("line" , xyToeSeam.x, xyO.y, xyToeSeam);

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


   if (   StringCompare("Straight", sToeStyle) || StringCompare ("Soft Enclosed B/M", sToeStyle) )
	AddEntity("Line", xyToeSeam, xyToeOFF) ;
   else {
	PRSetLayerNotes();
	AddEntity ("line", xyToeSeam.x, xyFirstTape.y, xyToePnt.x, xyFirstTape.y) ;
	}

// Foot lable
   PRSetLayerNotes() ;
   AddEntity ("text", sFootLabel, xyToeSeam.x + 1.5, xyToeSeam.y + 1) ;

// Seam Tramlines
// Add Closing lines at TOE
   AddEntity ("line" , xyO.x, xyO.y + nSeam, xyToeSeam.x, xyO.y + nSeam);
   AddEntity ("line" , xyO.x, xyO.y + nSeam + 0.5, xyToeSeam.x, xyO.y + nSeam + 0.5);
   }
else {
   // Panty
   PRSetLayerNotes() ;
   AddEntity ("line" , xyOtemplate.x, xyOtemplate.y + nSeam, xyO.x, xyOtemplate.y + nSeam);
   AddEntity ("line" , xyOtemplate.x, xyOtemplate.y + nSeam + 0.5, xyO.x, xyOtemplate.y + nSeam + 0.5);
   //Elastic note
   if (nElastic == 1 ) sFootLabel = "ELASTIC" ;
   if (nElastic == -1) sFootLabel = "NO ELASTIC" ;
   // Add the text as a symbol (For latter use by the StumpTool)
   if ( !Symbol("find",  "TextAsSymbol")) Display("message","error", "Elastic at distal end of support not added.\nCheck and add text manually");
   hEnt = AddEntity("symbol", "TextAsSymbol", xyLastTape.x + 0.25, xyOtemplate.y +  (xyLastTape.y - xyOtemplate.y)/2, 1, 1, 90 ) ;
   SetDBData(hEnt, "ID", sID+"PantyElasticNote") ;
   SetDBData(hEnt, "Data", sFootLabel ) ;
   
       // Panty Closing line	
   Execute ("menu", "SetLayer", hTemplateLayer);
   hEnt = AddEntity ("line", xyLastTape, xyOtemplate);
   SetDBData(hEnt, "ID", sID+"DistalClosingLine") ;
   AddEntity ("line", xyO, xyOtemplate);
   

   }
// Restore to layer 1

   Execute("menu", "SetLayer", Table("find", "layer", "1"));





 
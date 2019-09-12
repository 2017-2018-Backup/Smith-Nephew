// File Name:	LGLEGTMP.D
// Purpose:	Draws a lower extremity template 
//		from data held in a file
//		Draws leg points
//	
// Version:	1.01
// Date:		6.Jan94
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
//---------------------------------------------------------------------------------------------
// 04.Dec.94	GG	Draw foot scales only
//			For JOBSTEX Gradient 
//			Fabric Class
// 09.Dec.94	GG	Modified for Anklet, Knee and Thigh
//
// 20.Mar.94	GG	Patient details drawn as a single entity
//			using \n for new lines

// set template defaults
   nOffset = 0.5;

// Get Origin
   GetUser ("xy", "Give origin point for Leg", &xyOtemplate) ;
   Display ("cursor", "wait", "Drawing Leg");

// Set Symbol Library
   SetSymbolLibrary( sPathJOBST + "\\JOBST.SLB");

// Set layer   
   PRSetLayerConstruct () ;   

// Load template data file
   PROpenTemplateFile ();

   xyPt1 = xyOtemplate ;
   xyPt1.y = xyOtemplate.y + nOffset + nSeam ;

// Add a marker at the Origin for later use in drawing the foot 
   hEnt = AddEntity("marker","xmarker",xyOtemplate , 0.1, 0.1) ;	
   SetDBData (hEnt,"ID",sLegStyle + sFileNo + sLeg + "Origin" );
   SetDBData (hEnt,"units", sUnits) ;
   SetDBData (hEnt,"Data", sType) ;

// Loop Through data file to get to start of leg tapes
// Ignoring the none relevent ones 
nn = 1 ;
while (nn < nStyleFirstTape) {
	GetLine(hChan, &sLine) ;
	nn = nn  + 1 ;
	}
ScanLine( sLine, "blank", &nNo, &sScale, &nSpace, &n20Len, &nReduction);
nSpace = 0 ;

// Loop through displaying points and scales
//
if (nFabricClass == 2) nLoop = nAnkleTape - 1 ;  // JOBSTEX Gradient
	else nLoop = nStyleLastTape; 
while ( nn <=  nLoop ) {
   // Get data for leg and place on template
   if ( FNGetTape(nn)) {
	// Pleats, NB nSpace is distance of current tape from the previous tape
	nPleatGiven = 0 ;
	if ((nn == nStyleFirstTape+1) && (nFootPleat1 !=0)){ nSpace = nFootPleat1 ; nPleatGiven = 1;}
	if ((nn == nStyleFirstTape+2) && (nFootPleat2 !=0)){ nSpace = nFootPleat2; nPleatGiven = 1;}
	if ((nn == nStyleLastTape-1) && (nTopLegPleat2 !=0)){ nSpace = nTopLegPleat2; nPleatGiven = 1;}
	if ((nn == nStyleLastTape) && (nTopLegPleat1 !=0)){nSpace =  nTopLegPleat1 ; nPleatGiven = 1;}
	xyPt1.x = xyPt1.x + nSpace ;
	xyTmp.x = xyPt1.x ;
	nTapeLen = FNRound( FNDecimalise (FNGetTape(nn)) * nUnitsFac) ; 
	nLength = n20Len / 20 * nTapeLen   ; 
	xyTmp.y = xyOtemplate.y + nSeam + nLength ;
	AddEntity ("marker", "xmarker", xyTmp, 0.2, 0.2) ;

	// Add Pleat indication
	if (nPleatGiven == 1){
		SetData("TextHorzJust", 2);
		AddEntity ("text", "PLEAT", xyTmp.x - (nSpace / 2), xyTmp.y + 0.5 ) ;
		}

	// Add Ticks and Text at each scale
	// Start ticks 2 below and 2 above
	ii = int ( nLength / (n20Len / 20)) - 1 ;
                jj = ii + 2 ;

	nTickStep = n20Len / (20 * 8);
	nTickHt = xyOtemplate.y + nSeam + ii * nTickStep * 8;
	// Add Reduction Text
	SetData("TextHeight", 0.125);
	SetData("TextHorzJust", 1);
	AddEntity ("text", MakeString("long", nReduction), xyPt1.x + 0.2, nTickHt - 0.4) ;
	AddEntity ("text", Format("length", nTapeLen), xyPt1.x + 0.2, nTickHt - nSeam) ;
	// Print revised reduction at ankle tape except for anklets
	if (nn == nAnkleTape && nLegStyle != 0 ) {
		AddEntity ("text", sReductionAnkle, xyPt1.x + 0.4, nTickHt - 0.4) ;
		if (nFabricClass == 0) AddEntity ("text", sGramsAnkle + " grams", xyPt1.x + 0.4, nTickHt - 0.6) ;
		else  AddEntity ("text", sGramsAnkle + " stretch", xyPt1.x + 0.4, nTickHt - 0.6) ;
		AddEntity ("text", sMMAnkle + " mm", xyPt1.x + 0.4, nTickHt - 0.8) ;
		AddEntity ("marker", "xmarker", xyPt1.x + 0.3, nTickHt - 0.5, 0.2 , 0.2) ;
		}
	// Add tape ID
   	sSymbol = MakeString("long", nNo) + "tape" ;
   	if ( !Symbol("find", sSymbol)) Exit(%cancel, "Can't find a symbol to insert\nCheck your installation, that JOBST.SLB exists");
	AddEntity("symbol", sSymbol,  xyPt1.x, nTickHt ) ;
	while (ii <= jj ) {
		nTickLength = 0.22;
		AddEntity( "line", xyPt1.x, nTickHt, xyPt1.x + nTickLength, nTickHt) ;
		// Add Tick Text
		if ( ii < 10)  SetData("TextHorzJust", 4);
			else SetData("TextHorzJust", 2);
		AddEntity("text", MakeString("long",ii), xyPt1.x + nSeam, nTickHt - 0.01, 0.06, 0.1, 90 ) ; 
		nTickLength = 0.05;
		AddEntity( "line", xyPt1.x, nTickHt + nTickStep, xyPt1.x + nTickLength, nTickHt + nTickStep) ;
		AddEntity( "line", xyPt1.x, nTickHt + nTickStep*3, xyPt1.x + nTickLength, nTickHt + nTickStep*3) ;
		AddEntity( "line", xyPt1.x, nTickHt + nTickStep*5, xyPt1.x + nTickLength, nTickHt + nTickStep*5) ;
		AddEntity( "line", xyPt1.x, nTickHt + nTickStep*7, xyPt1.x + nTickLength, nTickHt + nTickStep*7) ;
		nTickLength = 0.08;
		AddEntity( "line", xyPt1.x, nTickHt + nTickStep*2, xyPt1.x + nTickLength, nTickHt + nTickStep*2) ;
		AddEntity( "line", xyPt1.x, nTickHt + nTickStep*6, xyPt1.x + nTickLength, nTickHt + nTickStep*6) ;
		nTickLength = 0.12;
		AddEntity( "line", xyPt1.x, nTickHt + nTickStep*4, xyPt1.x + nTickLength, nTickHt + nTickStep*4) ;
		ii = ii + 1;
		nTickHt = xyOtemplate.y + nSeam + ii * nTickStep * 8;
		}
	}
   nn = nn + 1 ;
   GetLine(hChan, &sLine) ;
   ScanLine( sLine, "blank", &nNo, &sScale, &nSpace, &n20Len, &nReduction);

// draw patient data at HEEL tape
// NOTE Heel tape is tape #7 in this case
   if ((nn == 7) || (FootLess && nn == nStyleFirstTape + 2)) {
	PRSetLayerNotes() ;
  	SetData("TextVertJust", 32);		// Bottom
   	SetData("TextHorzJust", 1);		// Left
   	SetData("TextHeight", 0.1);
	nTxtY = xyOtemplate.y + nSeam + 2.25 ;
	nTxtX = xyPt1.x ;	
     	if ( StringCompare(sLeg , "Left") ) sText = "Left Leg\n  " ;
	            else  sText = "Right Leg\n  ";
	// Data
  	sText = sText + sPatient  + "\n  " + sWorkOrder ;
	if (nFabricClass == 0)  sText = sText + "\n  " + StringMiddle(sFabric,5,StringLength(sFabric) - 4);
		else sText = sText + "\n  " +  sFabric;
 	AddEntity("text", sText, nTxtX, nTxtY);	
   	
   	PRSetLayerConstruct() ;
 	sText = "  " + sFileNo + "\n  "  + sDiagnosis + "\n  " + sAge + "\n  "  + sSEX  + "\n  "  + sPressure;
   	AddEntity("text", sText, nTxtX, nTxtY-0.75);	
	
	// Reset Setup text defaults 
   	SetData("TextFont", 0);
   	SetData("TextVertJust", 8);
   	SetData("TextHorzJust", 1);
   	SetData("TextHeight", 0.125);
   	SetData("TextAspect", 0.6);
	}
	
   } // End while

// xyO is the extreme right of drawing at this stage
   xyO.x = xyPt1.x ;
   xyO.y = xyOtemplate.y ;
   Close ("file", hChan) ;



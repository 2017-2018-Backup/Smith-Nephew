// File Name:	BODSLESS.D
// Purpose:		To draw the arm hole and shoulder line
//			for the sleeveless vest
// Version:		1.02
// Date:		16.Oct.95
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
//---------------------------------------------------------------------------------------------
// 16.Oct.97		gg 	Mods as G.Dunne Letter 14-10-97

//--------------------------------------------------------------------------------------------------------------------------------------
// --------------------------------------------------   SleeveLess Axilla ------------------------------------------
//--------------------------------------------------------------------------------------------------------------------------------------
// From the LowAxilla point traverse the VESTCURV picking up the following intersections
// in the order given (ie Increasing X)
// 
//	1.  xySleeveLess,
//

   if (StringCompare( sAxillaType, "Sleeveless")){
   	xyRaglanNeck = xyBackNeckConstruct ;
   	// Use 1/2 figured shoulder circumference 
   	// NB use of 1/2 scale
	if ( StringCompare ("Left", sSide))  nSLessFac  = FNRound( nLtSLessCir / 2);
		else nSLessFac  = FNRound(nRtSLessCir / 2);
	nSLessFac  = nSLessFac  - 0.5;	// 16.Oct.97 
	if (BRAGiven) {
		aAngle = Calc("angle", xyBreast, xyAxilla) ;
		xyAxillaLow = xyBreast ;	// Initail setting	 	
		}
	else	{
		aAngle = Calc("angle", xyWaistOFF, xyAxilla) ;
		xyAxillaLow = xyWaistOFF;	 // Initail setting	  		 	
		}
		
	// Revise xyAxillaLow
	xyAxillaLow.y = xyAxillaLow.y + ( tan(aAngle) *  abs(((nLowSLine + xyO.x) - nSLessFac ) - xyAxillaLow.x ) ) ; 	 	
	xyAxillaLow.x = nLowSLine - nSLessFac + xyO.x ;
		 	
	// Subtract for scoop necks
	if ( StringCompare("Scoop",sNeckType) || StringCompare("Measured Scoop",sNeckType))
		nSWidth = nSWidth - 1 ;	
	
	// Minimum width 1.5"
	if ( nSWidth < 1.5) nSWidth = 1.5 ;
	
  	nError = FN_CirLinInt( (nLowSLine + 0.125) + xyO.x , xyO.y, 
		(nLowSLine + 0.125)  + xyO.x, 100.0 + xyO.y, 
		xyRaglanNeck,
		nSWidth  ) ;
	xySleeveLess = xyInt ;
	
	// Shorten to accomodate body width
	if ( xySleeveLess.y > xyAxillaLow.y) xySleeveLess.y = xyAxillaLow.y; 

//   AddEntity ("marker", "xmarker", xySleeveLess , 0.25, 0.25, 0);
//   AddEntity ("marker", "xmarker", xyAxillaLow , 0.25, 0.25, 0);

	nAxillaShoulderRad =  Calc ("length", xyAxillaLow , xySleeveLess );  // V.I.P.
	aCurve = FN_CurveAngle ( xyAxillaLow , nAxillaShoulderRad ) ;
	aVestCurve =  Calc ("angle", xyAxillaLow , xySleeveLess);  	// V.I.P.	
	
	hCurve = Open ("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly");
   	GetLine(hCurve, &sLine);
   	ScanLine( sLine, "blank", &nLength, &aAngle);
	aPrevAngle = (aVestCurve - aCurve); 
    	xyPt1 = xyAxillaLow;
   	DrawSegment = %true ;					// Set draw segments flag on
   	while (GetLine(hCurve, &sLine)){
		aAngle = aAngle + aPrevAngle;
		xyPt2 = CalcXY("relpolar", xyPt1,nLength, aAngle);
		if ( FN_CirLinInt( xyPt1, xyPt2, xyAxillaLow, nAxillaShoulderRad )) {
			AddEntity("line", xyPt1, xyInt);
   			xySleeveLess = xyInt ;
			break;
			}
		if ( DrawSegment ) AddEntity("line",xyPt1,xyPt2);
		xyPt1 = xyPt2 ;
		aPrevAngle = aAngle;
		ScanLine( sLine, "blank", &nLength, &aAngle);
   	} //End GetLine while
          
          AddEntity("line", xySleeveLess, xyRaglanNeck);

          Close ("file", hCurve);
          

   } // Endif

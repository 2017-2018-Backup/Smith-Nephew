// File Name:	WHFTPNTS.D
// Purpose:	Establish Foot points
//
// Version:	1.03
// Date:		25.Jan.94
// Author:	Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
// 12.Mar.94	GG	Modifications wrt JOBST comments
//			Letter dated 4/3/94 parts 1,5
//			Fax dated 18/2/94  parts 2,3,7
// 28.Sept.94	GG	Modifications wrt Production Evaluation
//			Fax 9.Sept.94
// 09.Jan.95	GG	Use Long toe for both if the other leg -3
//			tape is > 10"
//

// Load template data file
   PROpenTemplateFile ();

// From data file retain reduction data of first scale for use with
// Straight toes (where a 14 reduction is required)
   nFirst20Len = n20Len ;
   nFirstReduction = nReduction ;
	
   nHeelTape = 6 ;
// Establish type of heel given the ankle position
// NB this comes via WH_FIGUR.D and the Data Base Field AnkleTape 
   if ( nAnkleTape == 7) {
	SmallHeel = %true ;
	nAnkleToHeelOffset = 1 ;
	}
   else	{
	SmallHeel = %false ;
	nAnkleToHeelOffset = 2 ;
	}

// Get data from both file and tapes for FOOT
// Actual position depends on ankle position ( this is based on heel size)
// 	1. First Tape
//	2. AnkleM ( -1.5 tape or -3 tape) (AnkleM stands for Ankle Minus)
//	2a. Previous tape to AnkleM ( -3 or - 4.5)
//	3. Heel 	( 0 tape)
//	4. Ankle ( +1.5 tape or +3 Tape)
    nHt =  xyOtemplate.x ;

// Loop Through data file to get to start of leg tapes
// Ignoring the none relevent ones 
   nn = 1 ;
   while (nn < nFirstTape) {
	GetLine(hChan, &sLine) ;
	nn = nn  + 1 ;
	}
    ScanLine( sLine, "blank", &nNo, &sScale, &nSpace, &n20Len, &nTmpltReduction);
    nSpace = 0 ;
    while ( nn <= nAnkleTape) {
	if ( (nn == nFirstTape + 1 ) && (nFootPleat1 != 0) ) nSpace =  nFootPleat1;
	if ( (nn == nFirstTape + 2) && (nFootPleat2 != 0) ) nSpace = nFootPleat2;
	nHt = nHt + nSpace ;
	if ( nn == nFirstTape) {
		xyFirstTape.x = nHt ;
		nLastTape20Len = n20Len;
		nLastTapeTmpltReduction = nTmpltReduction ;
		}	
	
	if ( nn == nHeelTape - nAnkleToHeelOffset) {
		xyAnkleM.x = nHt;
		nAnkleMTape = nHeelTape - nAnkleToHeelOffset ;
		nAnkleM20Len = n20Len ;
		nAnkleMTmpltReduction = nTmpltReduction ;
		if (nn == nFirstTape){
			nAnkleMPrev20Len = n20Len ;
			nAnkleMPrevTmpltReduction = nTmpltReduction ;
			nAnkleMPrevHt = nHt ;
			}
		else {	
			nAnkleMPrev20Len = nPrev20Len ;
			nAnkleMPrevTmpltReduction = nPrevTmpltReduction ;
			nAnkleMPrevHt = nPrevHt ;
			}
		}
	if  ( nn == nHeelTape) {
		xyHeel.x = nHt;
		nHeel20Len = n20Len ;
		nHeelTmpltReduction = nTmpltReduction ;
		}
	if ( nn == nAnkleTape) {
		xyAnkle.x = nHt ;
		nAnkle20Len = n20Len ;
		nAnkleTmpltReduction = nTmpltReduction ;
		}
	GetLine(hChan, &sLine);
	nPrevHt = nHt ;	// Store wrt AnklePrevM  & Foot Pleats
	nPrev20Len = n20Len ;
	nPrevTmpltReduction = nTmpltReduction ;
    	ScanLine( sLine, "blank", &nNo, &sScale, &nSpace, &n20Len, &nTmpltReduction);
	nn = nn + 1 ;
	}
   Close ("file", hChan) ;


// Foot Reductions Chart
// Quick and Dirty
   if ( nReductionAnkle >= 14 && nReductionAnkle <= 18) {
	nHeelChartReduction = 17 ;
	nAnkleMChartReduction = 14 ;
	}
   if ( nReductionAnkle >= 19 && nReductionAnkle <= 23) {
	nHeelChartReduction = 19 ;
	nAnkleMChartReduction = 17 ;
	}
   if ( nReductionAnkle >= 24 ) {
	nHeelChartReduction = 21 ;
	nAnkleMChartReduction = 17 ;
	}
   nAnkleMPrevChartReduction = 16 ;
   nToeCutBackReduction = 34 ;
   nToeCurvedReduction = 34 ;

// Choose heel plate
   if ( SmallHeel ) {
	// Heel plate #1
	nHeelR1 = 1.1760709;
	nHeelR2 = 0.8716718;
	nHeelR3 = 1.1760709; //Heel Plate now symetrical 27.Sept.94
	nHeelD1 = 2.0963720;
	nHeelD2 = 2.1843981;
  	nBackHeelOff = 0.05;
   	nFrontHeelOff = 0.1875;
	}
   else {
	// Heel plate #2
	nHeelR1 = 1.6601037 ;
	nHeelR2 = 1.0876233 ;
	nHeelR3 = 1.5053069 ;
	nHeelD1 = 2.7812341 ;
	nHeelD2 = 2.6294854 ;
  	nBackHeelOff = 0.25;
   	nFrontHeelOff = nBackHeelOff;
	}

// Calculate Control points
// For Heel
// At Ankle (ie +3 or + 1.5 Tape)

   nTapeLen = FNRound( FNDecimalise (FNGetTape(nAnkleTape)) * nUnitsFac) ; 
   if ( nFabricClass == 2 )
	nLength = (nTapeLen * ((100 -  nReductionAnkle) / 100)) /  2 ;
   else {
   	nLength = nAnkle20Len / 20 * nTapeLen   ;
   	nRedStep = nAnkle20Len /(20 * 8) ;
   	nLength = nLength + ((nAnkleTmpltReduction - nReductionAnkle) * nRedStep) ; 
	}	
   xyAnkle.y = xyOtemplate.y + nSeam + nLength ;

// At Heel (ie at 0 Tape)

   nTapeLen = FNRound( FNDecimalise (FNGetTape(nHeelTape)) * nUnitsFac) ; 
   nLength = nHeel20Len / 20 * nTapeLen   ;
   nRedStep = nHeel20Len /(20 * 8) ;
   nLength = nLength + ((nHeelTmpltReduction - nHeelChartReduction) * nRedStep) ; 
   xyHeel.y = xyOtemplate.y + nSeam + nLength ;

// At Minus Ankle ( ie -3 or -1.5 Tape)

   nAnkleMTapeLen = FNRound( FNDecimalise (FNGetTape(nAnkleMTape)) * nUnitsFac) ; 
   nLength = nAnkleM20Len / 20 * nAnkleMTapeLen   ;
   nRedStep = nAnkleM20Len /(20 * 8) ;
   nLength = nLength + ((nAnkleMTmpltReduction - nAnkleMChartReduction) * nRedStep) ;
   xyAnkleM.y = xyOtemplate.y + nSeam + nLength ;

   xyAnkleMPrev.x = nAnkleMPrevHt ;
   nLength = nAnkleMPrev20Len / 20 * nAnkleMTapeLen   ; // NB "nAnkleMTapeLen" from above
   nRedStep = nAnkleMPrev20Len /(20 * 8) ;
   nLength = nLength + ((nAnkleMPrevTmpltReduction - nAnkleMPrevChartReduction) * nRedStep) ; 
   xyAnkleMPrev.y = xyOtemplate.y + nSeam + nLength ;

// At Last tape using Minus Ankle Tape Value
// N.B. Quick and dirty
   nLength = nAnkleMTapeLen * 0.66 / 2 ;  // NB "nAnkleMTapeLen" from above
   xyFirstTape.y =  xyOtemplate.y  + nLength + nSeam ;

// Heel Circles
// Control points

   xyHeelCntrMidDistal.x = xyHeel.x - nFrontHeelOff ;
   xyHeelCntrMidDistal.y = xyHeel.y - nHeelR2 ;

   if (SmallHeel) {
	nCalc = Calc ("length", xyAnkleM, xyHeelCntrMidDistal);
	nAngle = Calc ("angle",xyHeelCntrMidDistal,xyAnkleM) - acos ((nHeelR1^2 - nHeelD1^2 - nCalc^2)/ (-2*nHeelD1*nCalc));
	xyHeelCntrDistal = CalcXY ("relpolar", xyHeelCntrMidDistal, nHeelD1, nAngle);
	}
   else {
   	nError = FN_CirLinInt( xyHeelCntrMidDistal.x,          xyAnkleM.y + nHeelR1 ,
		       	    xyHeelCntrMidDistal.x - 10,   xyAnkleM.y + nHeelR1 ,
		       	    xyHeelCntrMidDistal,
		                    nHeelD1) ;
  	xyHeelCntrDistal = xyInt ;
	}

   xyHeelCntrMidProximal.x = xyHeel.x + nBackHeelOff ;
   xyHeelCntrMidProximal.y = xyHeel.y - nHeelR2 ;
  
   BigAnkle = %false;	//Flag to indicate Ankle of Lymphdema proportions
   if (SmallHeel) {
	nCalc = Calc ("length", xyAnkle, xyHeelCntrMidProximal);
	nAngle = Calc ("angle",xyHeelCntrMidProximal,xyAnkle) + acos ((nHeelR3^2 - nHeelD2^2 - nCalc^2)/ (-2*nHeelD2*nCalc));
	xyHeelCntrProximal = CalcXY ("relpolar", xyHeelCntrMidProximal, nHeelD2, nAngle);
	}
   else {
   	nError = FN_CirLinInt( xyHeelCntrMidProximal.x,          xyAnkle.y + nHeelR3 ,
		       	    xyHeelCntrMidProximal.x + 10,  xyAnkle.y + nHeelR3 ,
		       	    xyHeelCntrMidProximal,
		      	    nHeelD2) ;
   	xyHeelCntrProximal = xyInt ;
	if (nError == %false) BigAnkle = %true ;	// IE no intersection found
	}

// Toes
// Quick and dirty ;
   xyToeSeam.y = xyOtemplate.y + nSeam ;
   sFootLabel = " " ;
   if ( nAge <= 10 ) {
	if ( nAge <= 2 ) nToeOffset =  2.75;
	if ( nAge == 3 ) nToeOffset = 3.00;	// Fax 9.Sept.94, Item2
	if ( nAge == 4 ) nToeOffset = 3.25;
	if ( nAge == 5 ) nToeOffset = 3.375;
	if ( nAge == 6 ) nToeOffset = 3.625;
	if ( nAge == 7 ) nToeOffset = 3.875;
	if ( nAge == 8 ) nToeOffset = 4.00;
	if ( nAge == 9 ) nToeOffset = 4.25;
	if ( nAge == 10 ) nToeOffset = 4.5;
	}

   if ( StringCompare ("Curved", sToeStyle)) {
	if (Male) xyToeSeam.x = xyHeel.x - 4.75;
		else xyToeSeam.x = xyHeel.x - 4.5;
	if ( nFirstTape <= 2 || nAnkleMTapeLen >= 10 || nOtherAnkleMTapeLen >= 10)
		//Figuring as a LONG Curved Toe
		xyToeSeam.x = xyHeel.x - 5.5;
	if ( nAge <= 10 ) 
		// Figuring as a EXTRA SHORT  Curved Toe"
		xyToeSeam.x = xyHeel.x - nToeOffset;
	}

    if (    StringCompare ("Cut-Back", sToeStyle)
         || StringCompare ("Soft Enclosed", sToeStyle) ) {
	if ( Male) xyToeSeam.x = xyHeel.x - 4.75;
		else xyToeSeam.x = xyHeel.x - 4.5;
	if ( nAnkleMTapeLen >= 10  || nOtherAnkleMTapeLen >= 10)
		xyToeSeam.x = xyHeel.x - 5.5;
	}

    if ( StringCompare ("Soft Enclosed", sToeStyle)) {
	sFootLabel = "CAP" ;
	if ( nAge <= 10) xyToeSeam.x = xyHeel.x - nToeOffset ;
	}

   if (    StringCompare("Straight", sToeStyle)
        || StringCompare ("Soft Enclosed B/M", sToeStyle) ) {
	xyToeSeam.x = xyHeel.x - 3.5;
   	nLength = nFirst20Len / 20 * nAnkleMTapeLen   ;
   	nRedStep = nFirst20Len /(20 * 8) ;
   	nLength = nLength + ((nFirstReduction - 14) * nRedStep) ; 
   	xyToeOFF.y = xyOtemplate.y + nSeam + nLength ;
	xyToeOFF.x = xyToeSeam.x ;
   	if ( StringCompare ("Soft Enclosed B/M", sToeStyle)) sFootLabel = "CAP" ;
	}


   if (      StringCompare ("Self Enclosed", sToeStyle) 
        ||   (StringCompare ("Soft Enclosed", sToeStyle) 
        && Value("scalar",sFootLength) ) ) {
	if (  StringCompare ("Self Enclosed", sToeStyle)) 
		sFootLabel = "ENCLOSED";
	else 
		sFootLabel = "CAP" ; 
	if (  nFootLength = FNRound( FNDecimalise (Value("scalar",sFootLength)) * nUnitsFac) ){
		if (SmallHeel) nFootLength = FNRound (nFootLength * 0.9) ;
		               else nFootLength = FNRound (nFootLength * 0.83) ;
		nLength = (nHeel20Len / 20 * 12.5) + xyOtemplate.y + nSeam ;
		if (  nLength < xyHeel.y) {
			nError = FN_CirLinInt (  xyOtemplate.x - 20,  xyOtemplate.y + nSeam,
			 	      xyHeel.x ,  xyOtemplate.y + nSeam,
			        	      xyHeel.x , nLength,
			        	      nFootLength) ;
			}
		else	
			nError = FN_CirLinInt (  xyOtemplate.x - 20,  xyOtemplate.y + nSeam,
			 	      xyHeel.x ,  xyOtemplate.y + nSeam,
			        	      xyHeel,
			        	      nFootLength) ;
		if (nError)
			xyToeSeam = xyInt ;
		else {
			Display ("message", "error", "Can't position Toe with given Foot length") ;
			Exit (%cancel, "Error forming foot") ;
			}
		}
	else {
		Display ("message", "error", "A Foot length is required for Self Enclosed Toes") ;
		Exit (%cancel, "No Foot length given") ;
		}
	}

// Toe Points to position toe arcs
// 

// Toe circle constants
// Quick and Dirty
   nToeCntrMidToCntrLowY  = 3.3359296 ;
   nToeCntrMidToCntrLowX  = 8.0276805 ;
   nToeCntrMidToCntrHighY = 3.1684888 ;
   nToeCntrMidToCntrHighX = 2.1960337 ;
   nToeMidR = 0.2528866 ;
   nToeLowR = 8.4403592 ;
   nToeHighR = 4.1080042 ;

// Establish Low Toe Arc center

   nLength = xyFirstTape.y - xyToeSeam.y  ; 	// Toe Point line to seam

   if ( nToeCntrMidToCntrLowY < nLength ) {
	// Simple case
	// End of toe curve is above seam line
	xyToeCntrLow.y = xyFirstTape.y - nToeCntrMidToCntrLowY ;
	xyToeCntrLow.x = xyToeSeam.x - nToeLowR ;
	}
   else {
	// More Complex case
	// End of toe curve is intersected by the seam line
	xyToeCntrLow.y = xyFirstTape.y - nToeCntrMidToCntrLowY ;
	nLength = nToeCntrMidToCntrLowY - nLength ;
	nLength = sqrt((nToeLowR * nToeLowR) - (nLength * nLength)) ;
	xyToeCntrLow.x = xyToeSeam.x - nLength ;
	}

// Having established TOE Low circle center the rest follows 
//
   xyToeCntrMid.y = xyToeCntrLow.y +  nToeCntrMidToCntrLowY ;
   xyToeCntrMid.x = xyToeCntrLow.x +  nToeCntrMidToCntrLowX ;
   xyToeCntrHigh.y = xyToeCntrMid.y -  nToeCntrMidToCntrHighY ;
   xyToeCntrHigh.x = xyToeCntrMid.x +  nToeCntrMidToCntrHighX ;

//AddEntity ("circle", xyToeCntrHigh, nToeHighR ) ;
//AddEntity ("circle", xyToeCntrMid, nToeMidR ) ;
//AddEntity ("circle", xyToeCntrLow, nToeLowR ) ;
//AddEntity ("circle",   xyHeelCntrDistal, nHeelR1) ;
//AddEntity ("circle",   xyHeelCntrMidDistal, nHeelR2) ;
//AddEntity ("circle",   xyHeelCntrProximal, nHeelR3) ;
//AddEntity ("circle",   xyHeelCntrMidProximal, nHeelR2) ;

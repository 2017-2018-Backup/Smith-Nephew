// File Name:	BODY.D
// Purpose:	Get Torso details acording to JOBST 
//		Prescription order form. 
//		Draw Body portion of Garment.
// Version:	1.11
// Date:		25.July.93
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
//---------------------------------------------------------------------------------------------
// 03.11.93	gg	Added Lining and Open Axilla
//			Added Bra Cups.
// 08.Nov.93	gg	Revised Axilla to Normal, Lrg Adult, Child
//
// 20.Nov.93	gg	Changed w.r.t JOBST letter 10.Nov.93.
//			Rounding added.
// 23.Nov.93	gg	Changes w.r.t JOBST Site Visit 23.Nov.93
//		gg	Mesh Axilla added
//
// 18.Jan.94	gg	Bug Fixes
//			Move variable declartion to BODYVBLS.D
//
// 14.Jun.94	gg	Vest Mods Ref: JUNE.94 (1)
// 
// 11.Jul.94	gg	Bug fix JL3.1
//
// 11.Sep.95	gg	Modifications wrt zipper/velcro factors
//
// 12.Oct.95	gg	BODY.D is now invoked from a Visual Basic Programme.
//			Modification made in support of this.
//			This is hybrid method as the DRAFIX MACRO code
//			has been proven in production and we are
//			making the minimum changes to this code
//
// 1.Nov.95	gg	Modifications w.r.t Implementation of sleeveless vest
//		
// 16.Oct.97	gg	Changes wrt Letter 14-10-97 by G.Dunne
//			Removed nSeamAllowance from shoulder to under breast							
// 17.Nov.97	gg	Don't drop neck at CL for turtlenecks
//								


// Start "HourGlass" cursor
    Display ("cursor", "wait", "Drawing Vest");

    UserSelection ("clear");
    UserSelection ("update");
    
    Execute ( "menu", "SetStyle", Table( "find",  "style", "bylayer"));
    Execute ( "menu", "SetColor", Table( "find",  "color", "bylayer"));

// Find JOBST installed directory
// Set path to macros
//
   STRING		sPathJOBST;
   sPathJOBST = GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;
   SetData("PathMacro",  sPathJOBST + "\\VEST" );

// Load variables
//
   @BODYVBLS.D;

// Load Functions and Procedures
//
   @BODYDEF.D;
	
// Load  body  data 
// Note this file is created by the Visual Basic Program
   @C:\JOBST\DRAW.D;
 
// Load conversion factors, references and defaults
   @BODYFACT.D;
   nBraAxillaHt = 0 ;

// Get Body Origin
   GetUser("xy", "Give Origin of Vest template", &xyO);

// Place a marker at the start point for later use.
// Get a UID and create the unique 4 character start to the ID code
// Note this is a bit dogey if the drawing contains more than 9999 entities
    PRSetLayerConstruct () ;
    hEnt = AddEntity("marker", "xmarker", xyO, 0.125);
    if (hEnt) {
    		sID=StringMiddle(MakeString("long", UID( "get",hEnt)), 1, 4) ;
    		while (StringLength(sID) < 4) sID = sID +  " ";
    		sID = sID + sFileNo + sSide ;
		SetDBData(hEnt, "ID" ,sID);
    		 }


   Display ("cursor", "wait", "Drawing Vest");
   
// Figure body Dimensions, Write to Log
   if ( nEOSCir != 0) {
	if ( nEOSCirUserFac != 0) nEOSCirFac = nEOSCirUserFac ;
	nEOSCir = FNRound (nEOSCir * nUnitsFac * nEOSCirFac / 4);
	}

   if ( nWaistCirUserFac != 0) nWaistCirFac = nWaistCirUserFac ;
   nWaistCir = FNRound(nWaistCir * nUnitsFac * nWaistCirFac / 4);
   
   BRAGiven = %false ;
   if ( nBreastCir != 0) {
	BRAGiven = %true ;
	if ( nBreastCirUserFac != 0) nBreastCirFac = nBreastCirUserFac ;
	nBreastCir = FNRound(nBreastCir * nUnitsFac * nBreastCirFac / 4);
	nBreastCirActual = FNRound(nBreastCirActual  * nUnitsFac) ;
	}

   if ( nS_Breast != 0) nS_Breast = FNRound(nS_Breast * nUnitsFac) ;

   if ( nNippleCir != 0) nNippleCir = FNRound(nNippleCir * nUnitsFac) ;

   nChestCirActual = FNRound(nChestCir * nUnitsFac) ;
   nChestCir = FNRound(nChestCir * nUnitsFac * nChestCirFac / 4) ;

// Shoulder width is relevent for Sleeveless only
   nSWidth = FNRound(nSWidth * nUnitsFac) ;
    
   nNeckGiven = nNeckCir * nUnitsFac ;  // Retain for later use
   nNeckCir = FNRound((( nNeckCir * nUnitsFac * nNeckFac_1) / nNeckFac_2) - nNeckFac_3) ;

	// Regular Neck
	if (StringCompare( "Regular", sNeckType)) {
		nNeckGiven = FNRound(nNeckGiven * nNeckFac_1);
		sNeckNotes = "Regular Neck " + Format ( "length", nNeckGiven ) ;
		}

	// Scoop neck & Measured Scoop neck
	if ( StringCompare("Scoop",sNeckType) || StringCompare("Measured Scoop",sNeckType)){
		nNeckCirRetained = nNeckCir ;	// Keep old value of neck to get radius
		// Recalculate to eliminate any rounding error carry through
		// Use NeckGiven
		nNeckCir = FNRound(((( nNeckGiven * nNeckFac_1) / nNeckFac_2) - nNeckFac_3) * nNeckFac_4) ;
		sNeckNotes = sNeckType + " Neck";
		}

	// Turtle Necks Both attached and detachable
	// NB. Size extraction is VERY format dependant,
           	if (StringCompare("Turtle", sNeckType, 6)) {
		nTurtleLength = 0;
		nNeckGiven = FNRound (nNeckGiven * nNeckFac_1);
		nTurtleWidth =  FNRound(nNeckDimension * nUnitsFac) ;
		if (nTurtleWidth == 0) {
			if ( nAge < 10)  nTurtleWidth = 1 ;	// Children under 10 yrs old
				else nTurtleWidth = 2 ;	// Adults or Children over 10 yrs old
			nTurtleWidth = nTurtleWidth * nUnitsFac ;	
			sNeckDimension = MakeString ("scalar", nTurtleWidth);
			SetDBData(hBody, "NeckDimension", sNeckDimension);  // update bodybox
			}

		if ( StringCompare("Turtle Detach",sNeckType, 13 ) )  /* ie Detachable*/
		nTurtleLength = nNeckGiven;		
		sNeckNotes = sNeckType + "\n"  + Format("length", nNeckGiven) 
				          + " x "  + Format ("length", nTurtleWidth) ;
		}


// Closures
// Check for bracups and that selected closure is compatible
// If front or back not given explicity then apply as per defaults in manual
// NB Format of sClosure is vital

	if ( BRAGiven ) {
		if (     StringCompare(sClosure, "Front Velcro") 
		     ||  StringCompare(sClosure, "Front Velcro (Reversed)") ) /* Front Velcro */
			sClosure = "Front Zip";

		if (StringLength(sClosure)==5 || StringLength(sClosure)==3 )  /* Front or Zip */
			sClosure = "Front Zip" ;

		if (StringLength(sClosure)==6) /* Velcro */
			sClosure = "Back Velcro" ;
		}
	else /* No Bra Cups */
		{
 		if (StringLength(sClosure)==4 || StringLength(sClosure)==6) /* Back or Velcro */
			sClosure = "Back Velcro" ;
		if (StringLength(sClosure)==3) /* Zip */
			sClosure = "Back Zip" ;
		}
		

// Check for different Left and Right Axilla Types

// Convert Lengths, Set lowest and Highest shoulder points

   if (nS_EOS != 0) {
	EOSGiven = %true;
	nS_EOS = FNRound(nS_EOS * nUnitsFac);
	nLowSLine = nS_EOS ;
	}
   else {
	EOSGiven = %false;
	nLowSLine = FNRound(nS_Waist  * nUnitsFac);
	}

   nHighSLine = nLowSLine + nHighSLineFac ; 
	

// Main Program
// Using the figured data a number of Keypoints are established which can then be used
// to create a polyline to represent the body template.
// If Required left and right axilla are detailed seperatly

// End of Support (If given) and Waist points.
   if (EOSGiven) {
	xyEOSCL.x = xyO.x;
	xyEOSCL.y = xyO.y;
	xyEOSOFF.x = xyO.x;
	xyEOSOFF.y = nEOSCir + nSeamAllowance + xyO.y ;
//	xyWaistOFF.x = nLowSLine + nSeamAllowance - FNRound((nS_Waist  * nUnitsFac)) + xyO.x;
	xyWaistOFF.x = nLowSLine  - FNRound((nS_Waist  * nUnitsFac)) + xyO.x;
	xyWaistOFF.y = nWaistCir + nSeamAllowance + xyO.y;
	}
   else {
	xyWaistCL.x = xyO.x;
	xyWaistCL.y = xyO.y;
	xyWaistOFF.x = xyO.x;
	xyWaistOFF.y = nWaistCir + nSeamAllowance +xyO.y;
	}

// Point at just under breast if BRA mesurements are given
    if (BRAGiven) {
	xyBreast.y = nBreastCir + nSeamAllowance + xyO.y ;
//	xyBreast.x = nLowSLine + nSeamAllowance - nS_Breast + xyO.x; // 16.Oct.97
	xyBreast.x = nLowSLine  - nS_Breast + xyO.x;
	}

// Set text attributes
   SetData("TextFont", 0);
   SetData("TextVertJust", 32);		// Top
   SetData("TextHorzJust", 4);		// Right
   SetData("TextHeight", 0.125);
   SetData("TextAspect", 0.6);

// Axilla and related points
// Calculate control points (NB. special case for axillas of different hieghts)
// Get front neck and Raglan intersection
// Note - This is point that controls the angle of the raglan curve
// 	The variable name "nNeckActualCir" is missleading as it is a calculated circumference 

                nLtSCir = nLtSCir * nUnitsFac;
                nRtSCir = nRtSCir * nUnitsFac;
	if ( abs(nLtSCir - nRtSCir) <= 1) {
		nAxilla = 1;
		// By default figure for left only, this will apply to both left and right 
		nLtSCir = (nLtSCir + nRtSCir) / 2;
		nLtSLessCir = FNRound(nLtSCir * 0.9);  // SleeveLess
		nRtSLessCir = nLtSLessCir ;		    // SleeveLess
   		nLtSCir = FNRound(nLtSCir / nSCirFac);
//		nRtSCir = nLtSCir;				
 		nRtSCir = nLtSCir - 0.5;			   // 16.Oct.97
   		xyHighestAxilla.x = nLowSLine - nRtSCir + xyO.x;
   		xyHighestAxilla.y = nChestCir + nSeamAllowance + xyO.y;
		xyLowestAxilla = xyHighestAxilla;
		sHighestAxilla = "None";
 		}
	else {
		nAxilla = 2;
		nLtSLessCir = FNRound(nLtSCir * 0.9) ; // SleeveLess
//   		nLtSCir = FNRound(nLtSCir / nSCirFac);	
   		nLtSCir = FNRound(nLtSCir / nSCirFac) - 0.5; // 16.Oct.97
		nRtSLessCir = FNRound(nRtSCir * 0.9); // SleeveLess
//   		nRtSCir = FNRound(nRtSCir / nSCirFac);	
   		nRtSCir = FNRound(nRtSCir / nSCirFac) - 0.5; // 16.Oct.97
   		xyHighestAxilla.x = nLowSLine - min( nRtSCir, nLtSCir) + xyO.x;
  		xyHighestAxilla.y = nChestCir + nSeamAllowance + xyO.y;
		if (BRAGiven)   aAngle = Calc ("angle", xyHighestAxilla, xyBreast);
		 	else aAngle = Calc ("angle", xyHighestAxilla, xyWaistOFF);
		xyLowestAxilla = CalcXY ("relpolar", xyHighestAxilla, abs (nRtSCir - nLtSCir), aAngle);
		if ( nRtSCir < nLtSCir) sHighestAxilla = "Right";
		        	else sHighestAxilla = "Left";
 		}

   if ( !StringCompare( sLtAxillaType, sRtAxillaType) ) nAxilla = 2;

// Establish if either axilla is SleeveLess
//
   SleeveLess=  %false;
   if ( StringCompare( "Sleeveless", sRtAxillaType) || StringCompare( "Sleeveless", sLtAxillaType) ) SleeveLess =  %true;
  
   nNeckActualCir = Calc("length", nHighSLine+xyO.x, nNeckCir+xyO.y , nLowSLine+xyO.x, xyO.y);  // GOP 01-02/18, 6.1

// Back Neck Construction point
   if   (FN_CirLinInt( 	nHighSLine + xyO.x , xyO.y, 
		nHighSLine + xyO.x, 100.0 + xyO.y, 
		nLowSLine + xyO.x , xyO.y, 
		nNeckActualCir ))
	xyBackNeckConstruct = xyInt;
   else	
	Display ("message", "error", "Can't form Back neck with this data!");

//Get front neck intesection
// 
//N.B. If either axilla is sleeveless then the entire back neck is constructed 
//        w.r.t the sleeveless axilla
   if   (FN_CirLinInt( 	nHighSLine - nNeckFrontFac +xyO.x , xyO.y,
		nHighSLine - nNeckFrontFac +xyO.x , 100.0 + xyO.y,
		nLowSLine + xyO.x , xyO.y,
		nNeckActualCir )) {
	xyFrontNeckOFF = xyInt ;
	// note mods w.r.t. sleveless vest
	if (SleeveLess){
		 xyFrontNeckOthers  = xyFrontNeckOFF ;
		 xyFrontNeckOFF  = xyBackNeckConstruct ;
		 }
	}
       else 
	Display ("message", "error", "Can't form Front neck with this data\nCheck JOB.LOG");

// Calculations for front neck  
// 
   if (nAge <= nChildUpperAge)
	nCLNeckDrop = 0.25;	// 16.Oct.97
   else
	nCLNeckDrop = 0.5;	// 16.Oct.97
	
// No CL Neck drop for any turtle neck, overide above	
   if (StringCompare("Turtle", sNeckType, 6)) nCLNeckDrop = 0;	// 17.Nov.97

   if (StringCompare("Measured Scoop", sNeckType))
	xyFrontNeckCL.x = nLowSLine - (nNeckCirRetained + nCLNeckDrop ) + xyO.x;
   else
	xyFrontNeckCL.x = nLowSLine - (nNeckActualCir + nCLNeckDrop ) + xyO.x;

   xyFrontNeckCL.y = xyO.y;
   xyFrontNeckCen.x = nLowSLine + xyO.x;
   xyFrontNeckCen.y = xyO.y;

// Calculation for Back Neck 
//
// Note:- mods for a back scooped neck ( Same Radius is used for both)
   xyBackNeckCL.x = nLowSLine + nSeamAllowance + xyO.x;
   xyBackNeckCL.y = xyO.y;
   nLength = Calc ( "length", xyBackNeckCL, xyBackNeckConstruct);
   aAngle = Calc ( "angle",  xyBackNeckCL, xyBackNeckConstruct);
   nBackNeckRadius =  (nLength/2) / cos(aAngle);
   
   if (   StringCompare("Regular",sBackNeckType) ) {	// Regular Back Neck - Arc centre
	xyBackNeckCL.x = nLowSLine + nSeamAllowance + xyO.x;
	xyBackNeckCen.x = xyBackNeckCL.x + nBackNeckRadius ;
   	xyBackNeckCen.y = xyO.y; 
	}
    if (    StringCompare("Scoop", sBackNeckType)
         || StringCompare("Measured Scoop",sBackNeckType)) {	
	if (nAge <= nChildUpperAge)
		nLength = nBackNeckScoopFacChild ;	
	else
		nLength = nBackNeckScoopFacAdult ;		
	if (StringCompare("Measured Scoop",sBackNeckType))
		nLength = nBackNeckDim * nUnitsFac;
	xyBackNeckCL.x = nLowSLine +  nSeamAllowance +  xyO.x - nLength;
	nLength = Calc ( "length", xyBackNeckCL, xyBackNeckConstruct);
  	aAngle = Calc ( "angle",  xyBackNeckCL, xyBackNeckConstruct);
	xyBackNeckConstruct_2.x = xyBackNeckCL.x + cos(aAngle) * nLength / 2 ;
	xyBackNeckConstruct_2.y = xyBackNeckCL.y + sin(aAngle) * nLength / 2 ;
	nLength = sqrt ( (nBackNeckRadius * nBackNeckRadius) - ((nLength * nLength)/4) ) ;
   	xyBackNeckCen = CalcXY ("relpolar", xyBackNeckConstruct_2, nLength, aAngle + 270);
	}
	
// Get the vestcurve angle
// For Highest axilla
//   nLowestAxillaFrontNeckRad = Calc ("length", xyLowestAxilla, xyFrontNeckOFF); // Approx. only
   nHighestAxillaFrontNeckRad = Calc ("length", xyHighestAxilla, xyFrontNeckOFF);  // V.I.P.
   aHighestVestCurve = Calc ("angle", xyHighestAxilla, xyFrontNeckOFF);  	// V.I.P.	
   aHighestCurveRotation = FN_CurveAngle ( xyHighestAxilla, nHighestAxillaFrontNeckRad ) ;
   nInitial_nAxilla = nAxilla ; 	// Retain this number w.r.t axilla labeling
//-------------------------------- Main While Loop ---------------------------------------
//
while ( nAxilla > 0) {   
   // N.B. Complications due to introduction of Sleevelees
   //         It was never great but this is  £^$&^%&*&*&!  awful!
   if ( nAxilla == 2) {
	// Set up side
	if ( StringCompare ("Right", sHighestAxilla) || StringCompare ("None", sHighestAxilla)) {
		sAxillaType = sRtAxillaType; 
		sSide = "Right";}
	else 	{
		sAxillaType = sLtAxillaType; 	
		sSide = "Left";}
	aCurve = aHighestCurveRotation;
	aVestCurve = aHighestVestCurve;
	nAxillaFrontNeckRad = nHighestAxillaFrontNeckRad;
	// Recalculate rotations etc if either axilla is sleeveless
	if (SleeveLess && !StringCompare ("Sleeveless", sAxillaType)) {
  		nHighestAxillaFrontNeckRad = Calc ("length", xyHighestAxilla, xyFrontNeckOthers);  // V.I.P.
		aCurve =  FN_CurveAngle ( xyHighestAxilla, nHighestAxillaFrontNeckRad ) ;
		aVestCurve =  Calc ("angle", xyHighestAxilla, xyFrontNeckOthers); 
		nAxillaFrontNeckRad = nHighestAxillaFrontNeckRad;
		}
	xyAxilla = xyHighestAxilla;
	}	
       else {
	if (StringCompare ("None", sHighestAxilla))  {
		aCurve = aHighestCurveRotation;
		aVestCurve = aHighestVestCurve;
		nAxillaFrontNeckRad = nHighestAxillaFrontNeckRad;
		xyAxilla = xyHighestAxilla;
		sSide = "Left";
		if (SleeveLess && !StringCompare ("Sleeveless", sLtAxillaType)) {
  			nHighestAxillaFrontNeckRad = Calc ("length", xyHighestAxilla, xyFrontNeckOthers);  // V.I.P.
			aCurve =  FN_CurveAngle ( xyHighestAxilla, nHighestAxillaFrontNeckRad ) ;
			aVestCurve =  Calc ("angle", xyHighestAxilla, xyFrontNeckOthers); 
			nAxillaFrontNeckRad = nHighestAxillaFrontNeckRad;
			}
		}
	else	{
		// xyRaglanNeck is established from the previous axilla
		// it ensures that both curves go to the back neck
		
		nAxillaFrontNeckRad =  Calc ("length", xyLowestAxilla, xyRaglanNeck);  // V.I.P.
		aCurve = FN_CurveAngle ( xyLowestAxilla, nAxillaFrontNeckRad ) ;
		aVestCurve =  Calc ("angle", xyLowestAxilla, xyRaglanNeck);  	// V.I.P.	
		xyAxilla = xyLowestAxilla;
		
		// Set up for Left or right Hand for lower axilla
		if ( StringCompare ("Right", sHighestAxilla)) {
			sSide = "Left";
			sAxillaType = sLtAxillaType; }
		else {	sSide = "Right";
			sAxillaType = sRtAxillaType; }
		if (SleeveLess && !StringCompare ("Sleeveless", sAxillaType)) {
  			nAxillaFrontNeckRad = Calc ("length", xyLowestAxilla, xyFrontNeckOthers);  // V.I.P.
			aCurve =  FN_CurveAngle ( xyLowestAxilla, nAxillaFrontNeckRad ) ;
			aVestCurve =  Calc ("angle", xyLowestAxilla, xyFrontNeckOthers); 
			}
		}
	}


// Set up layers and colours for Left or right Hand 
   if ( StringCompare ("Left", sSide))  {
	PRSetLayerTemplateLeft () ;
	sAxillaType = sLtAxillaType;
	}
   else	{
	sAxillaType = sRtAxillaType;
	PRSetLayerTemplateRight () ;
	}

// Find axilla closest to shoulder for bra cup placement
   nBraAxillaHt = max ( nBraAxillaHt, xyAxilla.x) ;

// Do axilla.
// The axilla MACROS are set by the visual basic programme.
// I.E. DRAW_1.D points to the relevant MACRO/S

   @C:\JOBST\DRAW_1.D;
   
// Store xyRaglanNeck for use in arm transfer
//
   xyRaglanNeckOriginal = xyRaglanNeck ;
   
// Draw Profile
// On seperate Left and Right Layers,  Draw EOS Left hand only
   if (StringCompare(sSide , "Left")){
   
      if (SleeveLess) xyRaglanNeck = xyFrontNeckOFF  ;  
      
// BackNeck
   // Draw for measured scoops
   //
      if (StringCompare("Measured Scoop", sBackNeckType)) {
	nX = xyRaglanNeck.x - xyBackNeckCL.x ;
	nY = xyRaglanNeck.y - xyBackNeckCL.y ;
	// Draw a fitted polyline using 1/3 and 1/2 rule
	hEnt = AddEntity( "poly"
		, "openfitted"
		, xyBackNeckCL
		, xyBackNeckCL.x + nX / 3,  xyBackNeckCL.y + nY / 2
		, xyRaglanNeck) ;
    	SetDBData(hEnt, "Data",sID);
    	SetDBData(hEnt, "curvetype","vestbackneck");
	Display ("message", "OKquestion",
		"The BACK Measured Scoop Neck has been drawn as a POLYLINE\n" +
		"Edit this line and make it OPEN FITTED, this will then be a smooth line\n" ) ; 
	}
   else {
   // Draw for Regular and Scoop back neck 
   	nLength = Calc ( "length", xyBackNeckCL, xyRaglanNeck);
   	aAngle = Calc ( "angle",  xyBackNeckCL, xyRaglanNeck);
   	nLength = (nLength/2) / cos(aAngle);
   	xyCen.x = xyBackNeckCL.x + nLength;
   	xyCen.y = xyO.y; 
   	aAngle = 2 * (180 - aAngle - 90);
   	aPrevAngle = 180.00 - aAngle ;
   	hEnt = AddEntity ( "arc", xyCen, nLength, aPrevAngle, aAngle);
    	SetDBData(hEnt, "Data",sID);
    	SetDBData(hEnt, "curvetype","vestbackneck");
	}
//
// End back neck section
	StartPoly("polyline") ;
	   // EOS
   	   if ( EOSGiven ){
		AddVertex(xyBackNeckCL) ;
		AddVertex(xyEOSCL) ;
		AddVertex(xyEOSOFF) ;	
		AddVertex(xyWaistOFF) ;
   		}
  	    else {
		AddVertex(xyBackNeckCL) ;
		AddVertex(xyWaistCL) ;
		AddVertex(xyWaistOFF) ;	
		}
	   // Under Breast to Waist
	   if ( BRAGiven )
		AddVertex(xyBreast) ;
	EndPoly();
          hEnt = UID ("find", UID ("getmax")) ;
    	SetDBData(hEnt, "Data",sID);
    	SetDBData(hEnt, "curvetype","vestprofile");

	// Detachable turtlenecks
	// Draw in box 3.0 x 3.0
	if (StringCompare("Turtle Detach",sNeckType, 13 )) {
		xyPt1.y = xyAxilla.y + 0.25 ;
		xyPt1.x = xyO.x ;
		
		hEnt = AddEntity("poly","polyline", xyPt1, xyPt1.x, xyPt1.y + 3,
			xyPt1.x, xyPt1.y + 3, xyPt1.x +3,  xyPt1.y +3,
			xyPt1.x + 3 ,  xyPt1.y +3,   xyPt1.x + 3, xyPt1.y,
			xyPt1.x + 3 , xyPt1.y, xyPt1,
			xyPt1, xyPt1.x, xyPt1.y + 3);
    		SetDBData(hEnt, "Data",sID);
		}	
	}


// Draw Axilla dependant profiles
// Lowest axilla point to next lowest point
    if ( BRAGiven )
	hEnt = AddEntity ("line", xyAxillaLow, xyBreast) ;
    else 
	hEnt = AddEntity ("line", xyAxillaLow, xyWaistOFF) ;
//    SetDBData(hEnt, "ID",sID);


// Axilla profile and axilla inserts
//
   if ( StringCompare( "Regular",sAxillaType, 7)){
	// Draw profile 
	hEnt = AddEntity ( "line", xyAxillaConstruct_2, xyRaglanAxilla); 
//	SetDBData(hEnt, "Data",sID);
	aPrevAngle = Calc ( "angle", xyAxilla, xyAxillaLow);
	aAngle =  Calc ( "angle",  xyAxilla, xyAxillaConstruct_2) - aPrevAngle ;
	hEnt = AddEntity ( "arc", xyAxilla, nRegAxillaFac_1, aPrevAngle, aAngle); 
//	SetDBData(hEnt, "Data",sID);

	// Draw insert
	xyInsertConstruct_4 = CalcXY ("relpolar", xyAxillaLow
					 , sqrt ( 2*(nSeamAllowance*nSeamAllowance))
					 , aPrevAngle + 45) ;
	aPrevAngle = Calc ( "angle", xyAxilla, xyInsertConstruct_4  );
	aAngle =  Calc ( "angle",  xyAxilla, xyAxillaConstruct_2) - aPrevAngle ;

	hEnt = AddEntity ( "arc", TransXY( xyAxilla, "trs", nXInsert, nYInsert)
		          , nRegAxillaFac_1 + nSeamAllowance, aPrevAngle, aAngle); 
//	SetDBData(hEnt, "Data",sID);
	hEnt = AddEntity ("line", TransXY(xyInsertConstruct_4, "trs", nXInsert, nYInsert)
		          , TransXY(xyInsertConstruct_3, "trs", nXInsert, nYInsert)) ;
//	SetDBData(hEnt, "Data",sID);
	hEnt = AddEntity ("line", TransXY(xyInsertConstruct_1, "trs", nXInsert, nYInsert)
		          , TransXY(xyInsertConstruct_2, "trs", nXInsert, nYInsert)) ;	
//	SetDBData(hEnt, "Data",sID);
	hEnt = AddEntity ("line", TransXY(xyInsertConstruct_2, "trs", nXInsert, nYInsert)
		          , TransXY(xyRaglanAxilla, "trs", nXInsert, nYInsert)) ;	
//	SetDBData(hEnt, "Data",sID);

	// Draw seam line and stamp 
	// Note recalculation because of drawing on the notes layer
	PRSetLayerNotes();
	hEnt = AddEntity ( "line", TransXY(xyAxillaConstruct_2, "trs", nXInsert, nYInsert)
		           , TransXY(xyRaglanAxilla, "trs", nXInsert, nYInsert)); 
//	SetDBData(hEnt, "Data",sID);
	aPrevAngle = Calc ( "angle", xyAxilla,xyInsertConstruct_4);
	aAngle =  Calc ( "angle",  xyAxilla, xyAxillaConstruct_2) - aPrevAngle ;
	hEnt = AddEntity ( "arc", TransXY( xyAxilla, "trs", nXInsert, nYInsert)
		          , nRegAxillaFac_1, aPrevAngle, aAngle); 
//	SetDBData(hEnt, "Data",sID);
	if (StringCompare (sSide, "Left")) nAType = 1;
		else nAType = 2;
	if ( nInitial_nAxilla == 1) nAType = 0 ;
	PRDataStamp ( TransXY(xyAxilla.x - 1, xyAxilla.y - 1, "trs", nXInsert, nYInsert), 10 + nAType) ;
	}
// Restore original raglan neck 
// Edits w.r.t sleeveless vest cause this
   xyRaglanNeck = xyRaglanNeckOriginal ;

// Update bodybox with data for drawing of sleeves
   if (!StringCompare(sAxillaType, "Sleeveless")) {
   	nAxillaBackNeckRad = Calc ( "length", xyAxilla, xyRaglanNeck) ;
   	nShoulderToBackRaglan = Calc ("length",  xyRaglanNeck, xyBackNeckConstruct);
   	if (StringCompare (sSide, "Right") && !StringCompare ("None", sHighestAxilla )){
		// These are for the RIGHT axilla data transfer
    		SetDBData ( hBody,  "AFNRadRight", MakeString("scalar", nAxillaFrontNeckRad)) ;
    		SetDBData ( hBody,  "ABNRadRight", MakeString("scalar",nAxillaBackNeckRad)) ;
    		SetDBData ( hBody,  "SBRaglanRight", MakeString("scalar",nShoulderToBackRaglan)) ;
		}
   	else 	{
		// These are for the LEFT or BOTH  axilla data transfer
		SetDBData ( hBody,  "AxillaFrontNeckRad", MakeString("scalar",nAxillaFrontNeckRad)) ;
		SetDBData ( hBody,  "AxillaBackNeckRad", MakeString("scalar",nAxillaBackNeckRad)) ;
		SetDBData ( hBody,  "ShoulderToBackRaglan", MakeString("scalar",nShoulderToBackRaglan)) ;
		}
	}

// loop to next Axilla
   nAxilla = nAxilla - 1 ;
       
   } //End Figure both axilla while loop

// Set up notes layer
   PRSetLayerNotes() ;

// Bra Cups
// DRAW_2.D created by visual basic
   if ( BRAGiven) {@C:\JOBST\DRAW_2.D;}
 
// If one side is sleevles then lable the other side
   if(SleeveLess) {
   	xyPt1 = CalcXY("relpolar", xyAxilla, Calc("length", xyAxilla, xyBackNeckCL) / 2 , Calc("angle", xyAxilla, xyBackNeckCL)) ;
   	if (!StringCompare(sRtAxillaType, "Sleeveless")) AddEntity("text", "Right", xyPt1) ;
   	if (!StringCompare(sLtAxillaType, "Sleeveless")) AddEntity("text", "Left",  xyPt1) ;
  	 } 

// Front neck
   nNeckDimension = nNeckDimension * nUnitsFac ;
   if (StringCompare("Measured Scoop", sNeckType)) {
	nX = xyFrontNeckOFF.x - xyFrontNeckCL.x + nNeckDimension;
	nY = xyFrontNeckOFF.y - xyFrontNeckCL.y ;
	// Draw a fitted polyline using 1/3 and 2/3 rule
	hEnt = AddEntity( "poly"
		, "openfitted"
		, xyFrontNeckCL.x - nNeckDimension,               xyFrontNeckCL.y
		, xyFrontNeckCL.x - nNeckDimension + nX / 3,  xyFrontNeckCL.y + nY*2/3
		, xyFrontNeckOFF) ;
    	SetDBData(hEnt, "Data",sID);
    	SetDBData(hEnt, "curvetype","vestfrontneck");

	Display ("message", "OKquestion",
		"The FRONT Measured Scoop Neck has been drawn as a POLYLINE\n" +
		"Edit this line and make it OPEN FITTED, this will then be a smooth line" ) ; 
	}
   else {
   	// Revise xyFrontNeckCen	
//   	nLength =  Calc("length", xyFrontNeckCL, xyFrontNeckOFF) / 2;
//  	aAngle = Calc("angle", xyFrontNeckCL, xyFrontNeckOFF);
//	nLength = nLength /  cos(aAngle) ;
// 	xyFrontNeckCen.X = xyFrontNeckCL.X + nLength ; 
// 	xyFrontNeckCen.Y = xyO.Y ;
   	nLength =  Calc("length", xyFrontNeckCL, xyFrontNeckOFF) / 2;
   	aAngle = Calc("angle", xyFrontNeckCL, xyFrontNeckOFF);
	xyFrontNeckCen = CalcXY("relpolar",xyFrontNeckCL, nLength, aAngle);
	aAngle = aAngle - 90 ;
	nNeckFac_5 = 1.3 ;
	nLength = sqrt (((nNeckActualCir * nNeckFac_5 ) * (nNeckActualCir * nNeckFac_5 )) - (nLength * nLength));
	xyFrontNeckCen = CalcXY("relpolar",xyFrontNeckCen,  nLength, aAngle);
      	aPrevAngle =  Calc("angle", xyFrontNeckCen, xyFrontNeckOFF);
   	aAngle = Calc("angle", xyFrontNeckCen, xyFrontNeckCL) - aPrevAngle;
   	hEnt = AddEntity ( "arc", xyFrontNeckCen, nNeckActualCir * nNeckFac_5  , aPrevAngle, aAngle);
   	SetDBData(hEnt, "Data",sID);
    	SetDBData(hEnt, "curvetype","vestfrontneck");

	}
// Neck Notes
//
   SetData("TextVertJust", 32);
   if ( StringCompare("Turtle Detach",sNeckType, 13 )){ /* ie Detachable*/
	SetData("TextHorzJust", 1);
	xyPt1.x = xyO.x + 0.25 ;
	xyPt1.y = xyAxilla.y + 0.1 ;
  	AddEntity("text", sNeckNotes, xyPt1.x, xyPt1.y + 0.4);
	AddEntity("text", sPatient, xyPt1.x, xyPt1.y + 1.0 );
	AddEntity("text", sWorkOrder, xyPt1.x, xyPt1.y + 0.8);
	AddEntity("text", sAge + " yrs", xyPt1.x , xyPt1.y + 0.6 );
	SetData("TextHorzJust", 4);
   	aAngle = 180 - Calc("angle", xyFrontNeckCen, xyFrontNeckOFF);
 	AddEntity("text", sNeckNotes , CalcXY("relpolar",xyFrontNeckCen, nNeckActualCir * nNeckFac_5 + 0.25, (180 - aAngle/2)));
	}
   else
	{
	SetData("TextHorzJust", 4) ;
   	aAngle = 180 - Calc("angle", xyFrontNeckCen, xyFrontNeckOFF) ;
 	AddEntity("text", sNeckNotes + "     ", CalcXY("relpolar",xyFrontNeckCen, nNeckActualCir * nNeckFac_5, (180 - aAngle/2)));
	}
	
// Closures
   SetData("TextHorzJust", 2);	
   
   // Closure length factors
   if ( StringCompare (sClosure, "Velcro") 
       ||
        StringCompare (sClosure, "Front Velcro") 
       ||
        StringCompare (sClosure, "Front Velcro (Reversed)")
       ||
        StringCompare (sClosure, "Back Velcro") 
       ||
        StringCompare (sClosure, "Back Velcro (Reversed)")
      ) 	 
      	nClosureAllowance = 0.375; 	// Velcro
    else
 	nClosureAllowance = 0.125; 	// Zippers only
   
   if( StringCompare( "Front", sClosure, 5)) 
	if ( StringCompare ( "Measured Scoop", sNeckType) )
		nLength =   xyFrontNeckCL.x - xyO.x - nClosureAllowance - nNeckDimension ;
	else
		nLength =  xyFrontNeckCL.x - xyO.x - nClosureAllowance ;
   else	
	nLength =  xyBackNeckCL.x - xyO.x - nClosureAllowance ;

   nLength = FNRound ( nLength) ;


//  Compensate for turtle necks
   if (StringCompare("Turtle", sNeckType, 6) ){

	if (StringCompare("Turtle Detach", sNeckType, 13)) 
		// Detachable necks
		nLength = FNRound( nLength / nZipFac) ;		
	else 
		//Ordinary and Fabric necks
		nLength = FNRound( (nLength + nTurtleWidth) / nZipFac) ;
	}
    else	
	nLength = FNRound( nLength / nZipFac) ;

   AddEntity ("text", 
	      sClosure + ", " + Format("length", nLength),
	      xyO.x + nLength /3,
	      xyO.y + 0.5);

// Add Stamps at waist
   PRDataStamp (xyWaistOFF.x +2, xyWaistOFF.y - 2, 1) ;	// Add a body stamp

// Draw construction lines for checking purpose
   PRSetLayerConstruct () ;
   AddEntity ("line", nHighSLine + xyO.x,  xyO.y, nHighSLine + xyO.x, xyAxilla.y);
   if ( BRAGiven) AddEntity ("line", xyBreast, xyBreast.x, xyO.y) ;
   if ( EOSGiven) AddEntity ("line", xyEOSOFF, xyEOSOFF.x ,xyAxilla.y) ;
   hEnt = AddEntity ("line", nLowSLine + xyO.x, xyO.y, nLowSLine + xyO.x, xyAxilla.y);       	//SetDBData(hEnt, "ID",sID);
   hEnt =  AddEntity ("line", xyWaistOFF ,  xyWaistOFF.x ,xyAxilla.y) ; 				//SetDBData(hEnt, "ID",sID);
   hEnt = AddEntity ("marker", "xmarker",xyBackNeckConstruct, 0.25, 0.25, 0);
   SetDBData(hEnt, "Data",sID);
   SetDBData(hEnt, "curvetype","vestbackneckconstruct");

// Use vestbody to carry the sID to the sleeve drawing
   SetDBData(hBody, "ID",sID);

// Fix required if Vest details change
   if (StringCompare ("None", sHighestAxilla ) ) { 
    	SetDBData ( hBody,  "AFNRadRight", "") ;
    	SetDBData ( hBody,  "ABNRadRight", "") ;
    	SetDBData ( hBody,  "SBRaglanRight", "") ;
   	hEnt = AddEntity ("line", xyAxilla,  xyAxilla.x, xyO.y) ;
    	SetDBData(hEnt, "Data",sID);
    	SetDBData(hEnt, "curvetype","vest" + sSide + "axillaconstruct");
	}
   else {
   	hEnt = AddEntity ("line", xyHighestAxilla,  xyHighestAxilla.x, xyO.y) ;
    	SetDBData(hEnt, "Data",sID);
    	SetDBData(hEnt, "curvetype","vest" + sHighestAxilla + "axillaconstruct");
   	hEnt = AddEntity ("line", xyLowestAxilla,  xyLowestAxilla.x, xyO.y) ;
    	SetDBData(hEnt, "Data",sID);
    	if (StringCompare(sHighestAxilla, "Left"))
    		SetDBData(hEnt, "curvetype","vestRightaxillaconstruct");
    	else
    		SetDBData(hEnt, "curvetype","vestLeftaxillaconstruct");
    	if ( !SleeveLess) {
		Display ("message", "OKquestion",
		"WARNING!\n"+
		"You might need to transfer the lowest axilla front neck\n" +
		"to the sleeve raglan.\n" +
		"Use the \"Vest to Slv.\" tool\n " ) ; 
		}
	}
	
// Do Mesh axilla.
// The axilla MACROS are set by the visual basic programme.
// I.E. DRAW_1.D points to the relevant MACRO/S

 if (bPrevAxillaWasMesh ){  @C:\JOBST\DRAW_3.D;}

// End of BODY.D, Clean up and close down

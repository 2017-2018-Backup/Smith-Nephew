// File Name:	WH_STUMP.D
// Purpose:	Draws a stump
// Version:	1.03
// Date:		04.Jan.95
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
// 17.Sep.96	GG	Added code w.r.t using SetVertex on a curve
//			with only 3 vertex.
// 7.Oct.96	GG	Added code w.r.t CHAP

   Display ("cursor", "wait", "Loading");

HANDLE
	hCurv,
	hEnt,
	hChan,
	hLeg,
	hTitle,
	hClosingLine,
	hElasticNote,
	hLayer;

XY
	xyInt,	// Returned by FN_CirLinInt
	xyDistal,
	xyOtemplate,
	xyFirstProfile,
	xyPt1,
	xyPt2,
	xyPt3,
	xyApexPt,
	xyMidPt,
	xyTmp;

STRING
	sProfileID,
	sStyleID,
	sStyleString,
	sOtemplateID,
	sClosingLineID,
	sElasticNoteID,
	sUnits,
	sTmp,
	sPatient, sWorkOrder,
	sLeg,
	sTapeLengths,
	sFileNo,
	sType;

BOOL
	OriginFound;

ANGLE
	aAngle;
	
NUMBER
	ii,
	nn, nFirstTape, nLastTape,
	nAge ;
	
// Find JOBST installed directory
// Set path to macros
//
   STRING	 sPathJOBST;
   sPathJOBST = GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;
   SetData("PathMacro",  sPathJOBST + "\\LEG" );
 
// Sub Routines etc
//
   @LGSTMDEF.D;

// Reset 
    UserSelection ("clear");
    Execute ( "menu", "SetStyle", Table( "find",  "style", "bylayer"));
    Execute ( "menu", "SetColor", Table( "find",  "color", "bylayer"));

// Get profile, identify Leg and FileNo

   if(!hCurv = GetEntity ("Select a Leg Profile", "curve")) Exit (%ok, "User cancelled") ;

   Display ("cursor", "wait", "Searching for more data");
   GetDBData (hCurv, "ID", &sProfileID);
   nStringLength = StringLength(sProfileID);
   if (StringCompare( "LeftLegCurve", StringMiddle(sProfileID, nStringLength - 11, 12))){
	sStyleID = StringMiddle(sProfileID, 1, nStringLength - 12);
	sLeg = "Left";
	}
    if (StringCompare( "RightLegCurve", StringMiddle(sProfileID, nStringLength - 12, 13))) {
	sStyleID = StringMiddle(sProfileID, 1, nStringLength - 13);
	sLeg = "Right";
	}

// Check if this is a valid Leg Curve exit if not
    if (StringLength (sLeg) == 0) Exit (%cancel, "A Leg Profile was not selected\n") ;

// Get Marker data
   sStyleID = sStyleID + sLeg;
   sOtemplateID = sStyleID + "Origin";
   sClosingLineID = sStyleID + "DistalClosingLine";
   sElasticNoteID = sStyleID +  "PantyElasticNote";
   OriginFound = %false ;
   sTmp = "(type = 'Marker' AND  DB ID ='" + sOtemplateID + "') OR (type ='Symbol' AND DB ID ='" + sElasticNoteID + "') OR  (type ='line' AND DB ID ='" + sClosingLineID +  "')" ;
   hChan = Open ("selection", sTmp );
   if (hChan) {
	ResetSelection(hChan);
	while (hEnt = GetNextSelection(hChan)) {
		GetDBData(hEnt,"ID", &sTmp) ;
			if (StringCompare(sTmp, sOtemplateID)) {
				OriginFound = %true ;
				GetGeometry(hEnt, &sTmp, &xyOtemplate , &xyTmp, &aAngle) ;
  				GetDBData( hEnt,"TapeLengthsPt1", &sTapeLengths);
  				GetDBData( hEnt,"TapeLengthsPt2", &sTmp);
					sTapeLengths = sTapeLengths + sTmp ;
   				GetDBData ( hEnt,"FirstTape", &sTmp) ;
   				 	nFirstTape = Value ("scalar", sTmp);
      				GetDBData ( hEnt,"LastTape", &sTmp);
    				 	nLastTape= Value ("scalar", sTmp);
     				GetDBData ( hEnt, "age", &sTmp) ; 
   				 	nAge = Value ("scalar", sTmp);
   				GetDBData( hEnt, "units", &sUnits);
				// Get patient details for use in labeling circular stumps
				GetDBData( hEnt, "patient", &sPatient);
				GetDBData( hEnt, "WorkOrder", &sWorkOrder);
				}
			if (StringCompare(sTmp, sClosingLineID)) {
				hClosingLine = hEnt ;
				}
			if (StringCompare(sTmp, sElasticNoteID)) {
				hElasticNote = hEnt ;
				}
		}
   	}
   Close("selection", hChan); 

// Check if that the markers have been found, otherwise exit
//
   if (!OriginFound) Exit (%cancel, "Origin Marker not found!\nUnable to continue.") ;
    
// Set up units conversion factor
   if ( StringCompare(sUnits, "cm")) nUnitsFac = 10/25.4;	// Cm to Inches   
   if ( StringCompare(sUnits, "inches")) nUnitsFac = 1;	// Inches

// Release nFirstTape to a .78 reduction
    nTapeLen = FNRound(FNDecimalise(FNGetTape(nFirstTape)) * nUnitsFac) ;
    nLength = ((nTapeLen * 0.78) / 2) + nSeam;
    nTol = abs (nLength - ((nTapeLen * 0.95) + nSeam)) + .125;

    xyDistal.x = xyOtemplate.x ;
    xyDistal.y = xyOtemplate.y + nLength ;

// Do a check to ensure that first vertex of the profile and the point above are close
   GetVertex (hCurv, 1, &xyTmp) ;
   if (abs(xyDistal.x - xyTmp.x) > .125)
	Exit (%cancel, "Origin marker and Profile start not within tolerance in X !\n") ;
   if (abs(xyDistal.y - xyTmp.y) > nTol)
	Exit (%cancel, "Origin marker and Profile start not within tolerance in Y !\n") ;

// Draw Stumps
   xyFirstProfile = xyDistal;
   if (nAge <= 10) xyDistal.x = xyDistal.x - 0.5 ;
	else xyDistal.x = xyDistal.x - 0.75 ;

   if ( StringCompare(sLeg, "Left"))
	PRSetLayerTemplateLeft();
   else
	PRSetLayerTemplateRight();
 
   nLength = Calc ("length", xyOtemplate, xyDistal) ;
   aAngle =  Calc ("angle", xyOtemplate, xyDistal) ;

   if ( nLength > 3 ) {
	// STAR stump
      StartPoly("polyline");
	// First Point	
	nA = sqrt (((nLength/3)+0.125)^2 - (nLength/6)^2) ;
	xyPt2 = CalcXY("relpolar", xyOtemplate,  nLength/3 , aAngle) ;
	xyMidPt = CalcXY("relpolar", xyOtemplate,  nLength/6, aAngle) ;		
	xyApexPt = CalcXY("relpolar", xyMidPt,  nA, aAngle + 90) ;

	nError = FN_CirLinInt (xyApexPt, xyOtemplate, xyApexPt, 0.125) ;
	AddVertex( xyOtemplate);
	AddVertex( xyInt);

	nError = FN_CirLinInt (xyApexPt, xyPt2, xyApexPt, 0.125) ;
	AddVertex( xyInt);
	AddVertex( xyPt2);
	
	// Middle Point
	nA = sqrt (((nLength/3)+0.25)^2 - (nLength/6)^2) ;
	xyPt3 = CalcXY("relpolar", xyOtemplate,  (nLength/3)*2 , aAngle) ;
	xyMidPt = CalcXY("relpolar", xyOtemplate,  nLength/2, aAngle) ;		
	xyApexPt = CalcXY("relpolar", xyMidPt,  nA, aAngle + 90) ;

	nError = FN_CirLinInt (xyApexPt, xyPt2, xyApexPt, 0.25) ;
	AddVertex( xyInt);

	nError = FN_CirLinInt (xyApexPt, xyPt3, xyApexPt, 0.25) ;
	AddVertex( xyInt);
	AddVertex( xyPt3);

	// Last Point	
	nA = sqrt (((nLength/3)+0.125)^2 - (nLength/6)^2) ;
	xyMidPt = CalcXY("relpolar", xyOtemplate,  (nLength/6)*5, aAngle) ;		
	xyApexPt = CalcXY("relpolar", xyMidPt,  nA, aAngle + 90) ;

	nError = FN_CirLinInt (xyApexPt, xyPt3, xyApexPt, 0.125) ;
	AddVertex( xyInt);

	nError = FN_CirLinInt (xyApexPt, xyDistal, xyApexPt, 0.125) ;
	AddVertex( xyInt);
	AddVertex( xyDistal);
	AddVertex( xyFirstProfile );
       EndPoly() ;

	//Adjust profile
	//There is a DRAFIX bug that causes drafix to fail if there are only 3
	//vertex in the curve when using SetVertex, hence:-
	if (GetVertexCount( hCurv) >3) 
 		SetVertex (hCurv, 1, xyFirstProfile ) ;
 	else
 		Display("message", "error", "As the Curve only has 3 points! it can't be adjusted automatically.  You will need to move the end point to the top of the stump manually");

	// Ticks and lines
	PRSetLayerNotes();

	xyTmp = CalcXY("relpolar", xyPt2,  0.25, aAngle - 90) ;
	AddEntity("line", xyPt2, xyTmp);
	AddEntity("line",CalcXY("relpolar", xyTmp,  0.125, aAngle) 
		         ,CalcXY("relpolar", xyTmp,  0.125, aAngle+180));

	xyTmp = CalcXY("relpolar", xyPt3,  0.25, aAngle - 90) ;
	AddEntity("line", xyPt3, xyTmp);
	AddEntity("line",CalcXY("relpolar", xyTmp,  0.125, aAngle) 
		         ,CalcXY("relpolar", xyTmp,  0.125, aAngle+180));

  	AddEntity("marker","closed arrow", xyDistal.x + 0.5 , xyDistal.y,  0.5 , 0.125, 270 ) ;	

	}
   else {
	// Circular stump
	StartPoly ( "polyline" ) ;
	  AddVertex ( xyOtemplate );
	  AddVertex ( xyDistal );
	  AddVertex ( xyFirstProfile );
	EndPoly () ;

	//Adjust profile
	//There is a DRAFIX bug that causes drafix to fail if there are only 3
	//vertex in the curve when using SetVertex, hence:-
	if (GetVertexCount( hCurv) >3) 
 		SetVertex (hCurv, 1, xyFirstProfile ) ;
 	else
 		Display("message", "error", "As the Curve only has 3 points! it can't be adjusted automatically.  You will need to move the end point to the top of the stump manually");

	// Circle
	nLength = Calc("length", xyOtemplate ,xyDistal ) - nSeam ;
	nLength = nLength / 3.1416 ;
	xyTmp.y = xyOtemplate.y + nLength + 0.125;
	xyTmp.x = xyDistal.x - ( nLength + 0.125);

	AddEntity("circle", xyTmp, nLength + 0.125) ;

	PRSetLayerNotes();
	AddEntity("circle", xyTmp, nLength) ;
	
	SetData("TextHorzJust", 2) ;
	SetData("TextVertJust", 8) ;
	AddEntity("text", sPatient , xyTmp);

	SetData("TextVertJust", 32) ;
	AddEntity("text", sWorkOrder , xyTmp);
	}

// Clean up
   if (hElasticNote) DeleteEntity(hElasticNote) ;
   if (hClosingLine) DeleteEntity(hClosingLine) ;
//   if (hClosingLine) SetGeometry (hClosingLine, xyOtemplate, xyFirstProfile);

// Reset and exit
   Execute ("menu", "SetLayer",  Table("find", "layer", "1")) ;
   Exit (%ok, "Stump drawing Complete");

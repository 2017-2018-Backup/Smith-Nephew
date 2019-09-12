// File Name:	FIXREGLR.D
// Purpose:	TOOL to make the arc of the Regular 
//		Axilla cut-out touch the profile
//		Used after editing the profile.
//
// Version:	1.01
// Date:		16.May.94
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date	By	Action

XY	xyAxilla,	
	xyPt1, 
	xyPt2, 
	xyStart, 
	xyEnd;
STRING	sProfileID, 
	sArcID,
	sArm, 
	sEntClass, 
	sStyle,
	sTmp,
	sInsertArcID,
	sInsertSeamArcID,
	sInsertLineID;
ANGLE	aAngle, 
	aAngleChange,
	aStartAngle, 
	aDeltaAngle,
	aNewStartAngle, 
	aNewDeltaAngle;
HANDLE	hArc, 
	hCurv, 
	hInsertArc, 
	hInsertSeamArc, 
	hInsertLine, 
	hEnt, 
	hChan ;
NUMBER nAxillaRadius,
	nInsertBits, 
	nTape;
BOOL
	Change;

   UserSelection("clear") ;
   UserSelection("update") ;
   Execute ( "menu", "SetStyle", Table( "find",  "style", "bylayer"));
   Execute ( "menu", "SetColor", Table( "find",  "color", "bylayer"));

   hCurv = GetEntity ("Select a Arm Profile", "curve line polyline") ;
   if (!hCurv) Exit(%ok,"No Arm Profile selected");
   
   Display ("cursor", "wait", "Adjusting Axilla insert");
   
   GetDBData (hCurv, "ID", &sProfileID);
   nStringLength = StringLength(sProfileID);
   if (StringCompare( "LeftProfile", StringMiddle(sProfileID, nStringLength - 10, 11))){
	sStyle = StringMiddle(sProfileID, 1, nStringLength - 11);
	sArm = "Left";
	}
    if (StringCompare( "RightProfile",StringMiddle(sProfileID, nStringLength - 11, 12))) {
	sStyle = StringMiddle(sProfileID, 1, nStringLength - 12);
	sArm = "Right";
	}
	
// Check if this is a valid Arm Curve exit if not (otherwise it fails badly)
//
    if (StringLength (sArm) == 0) Exit (%cancel, "An ARM Profile was not selected\n") ;

   
// Get the insert
// Arc and line
//
   sArcID =  sStyle + sArm + "RegularAxillaArc";
   sInsertArcID   = sStyle + sArm + "RegularAxillaInsertArc";
   sInsertSeamArcID   = sStyle + sArm + "RegularAxillaInsertSeamArc";
   sInsertLineID = sStyle + sArm + "RegularAxillaInsertLine";
   sTmp = "DB ID = '" + sInsertSeamArcID + "' OR DB ID = '" + sArcID + "' OR DB ID = '" + sInsertArcID + "' OR DB ID ='"  + sInsertLineID + "'";
   nInsertBits = 0 ;
   hChan = Open ("selection", sTmp );
   if (hChan) {
	ResetSelection(hChan);
	while (hEnt = GetNextSelection(hChan)) {
		GetDBData(hEnt,"ID", &sTmp) ;
		if (StringCompare(sTmp, sArcID )){
			nInsertBits = nInsertBits + 1 ;
			hArc = hEnt;
			}
		if (StringCompare(sTmp, sInsertArcID )){
			nInsertBits = nInsertBits + 1 ;
			hInsertArc = hEnt;
			}
		if (StringCompare(sTmp, sInsertSeamArcID )){
			nInsertBits = nInsertBits + 1 ;
			hInsertSeamArc = hEnt;
			}
		if (StringCompare(sTmp, sInsertLineID )) {
			nInsertBits = nInsertBits + 1 ;
			hInsertLine = hEnt;
			}
		}
   	}
   Close("selection", hChan); 
 
// Check if that the Inserts have been found, otherwise exit
//
    if (nInsertBits < 4) Exit (%cancel, "An Axilla Inserts does not exist for this sleeve or \nelements of the insert are missing!\n") ;
    if (nInsertBits > 4) Exit (%cancel, "Two or more drawings of the same style exist!\nDelete the extra drawing/s and try again.\n") ;

//
// Function:  	number = FN_CirLinInt( xyLineStart, xyLineEnd, xyCircleCen, nRadius);
//
// Parameters:	xyLineStart = %1, %2
//		xyLineEnd = %3, %4
//		xyCirleCen = %5, %6
//		Radius = %7
//
// Returns:	%true if intersection found
//		%false if no intesection
//		xyInt = First intersection ( Global Variable )
//
//
// Variables:
XY	
	xyInt; 	// if it exist, 1st intersection of line and circle
NUMBER 
	nK, nA, nB, nC, nM, nCalcTmp, nRoot, nSlope, nSign;

function FN_CirLinInt()
{
//Note - Returns only the First intersection

   nSlope = Calc("angle", %1, %2, %3, %4);

// Horizontal Line
   if ( nSlope == 0 || nSlope == 180 ) {
	nSlope = -1;
	nC = %7^2 - (%2-%6)^2;
   	if ( nC < 0 ) return(%false);		// No Roots
	nSign = 1;			// Test each root
                while ( nSign > -2){
		nRoot = %5 + sqrt(nC) * nSign;		
  		if ( nRoot >= min(%1,%3) && nRoot <= max(%1,%3))  {
			xyInt.x = nRoot;
			xyInt.y = %2;
			return (%true);
			}
                	nSign = nSign - 2;
		}
	return(%false);
	}

// Vertical line
   if ( nSlope==90 || nSlope == 270 ) {
	nSlope = -1;	
	nC = %7^2 - (%1-%5)^2;
	if ( nC <0 ) return(%false);		// No Roots
	nSign = 1;			// Test each root
	while ( nSign > -2){
		nRoot = %6 + sqrt(nC) * nSign;
  		if ( nRoot >= min(%2,%4) && nRoot <= max(%2,%4)){
			xyInt.y = nRoot;
			xyInt.x = %1;
			return (%true);
			}
		nSign = nSign -2;
		}
	return(%false);
	}

// Non-Orthogonal Line
   if ( nSlope > 0 ) {
   	nM = (%4 - %2) / (%3 - %1) ;		// Slope
   	nK = %2 - nM*%1;			// Y-Axis intercept
   	nA = (1 + nM^2) ;			
   	nB = 2 * ( -%5 + nM*nK - %6*nM) ;
  	nC = %5^2 + nK^2 + %6^2 - 2*%6*nK - %7^2 ;
   	nCalcTmp = (nB^2) - (4 * nC *nA);

   	if ( nCalcTmp < 0 ) return (%false);	// No Roots
	nSign = 1;
	while ( nSign > -2) {
   		nRoot = ( -nB + (sqrt(nCalcTmp)/nSign))/(2*nA);
  		if ( nRoot >= min(%1,%3) && nRoot <= max(%1,%3)) {  
			xyInt.x = nRoot;
	   		xyInt.y = nM * nRoot + nK;
			return (%true);
			}
		nSign = nSign -2;
		}
	return(%false);
 	}

   } // End function FN_CirLinInt
		 

// Get data from Curve and Arc
//
   GetEntityClass(hCurv, &sEntClass) ;
   GetGeometry (hArc, &xyAxilla,  &nAxillaRadius, &aStartAngle, &aDeltaAngle);
  
// Find point on sleeve profile (note - this is an aproximation only)
// For curve profiles only, not short sleeves 3 tapes or less
   Change = %false ;
   if (StringCompare("curve", sEntClass) || StringCompare("polyline", sEntClass)) {
    	nTape = GetVertexCount(hCurv) - 1  ;  
    	xyPt1 = xyAxilla ;
   	while (nTape > 0 ) {
 		GetVertex( hCurv, nTape, &xyPt2) ;
		if (FN_CirLinInt ( xyPt1, xyPt2, xyAxilla,  nAxillaRadius)) {
        			aNewStartAngle = Calc ( "angle", xyAxilla, xyInt);
        			aAngleChange =  (aStartAngle - aNewStartAngle) ;
        			aNewDeltaAngle = aDeltaAngle + aAngleChange  ;
        			if (aNewStartAngle != aStartAngle) Change = %true ;		
			break ;	
			}
    	 	nTape = nTape - 1;	
		xyPt1 = xyPt2 ;
		}
	}
    else {
    	// As we know this is a "line" we can use GetExtent to get the start and endpoint
    	GetExtent(hCurv, &xyStart, &xyEnd) ;
	if (FN_CirLinInt ( xyStart, xyEnd, xyAxilla,  nAxillaRadius)) {
        		aNewStartAngle = Calc ( "angle", xyAxilla, xyInt);
        		aAngleChange =  (aStartAngle - aNewStartAngle) ;
        		aNewDeltaAngle = aDeltaAngle + aAngleChange  ;
        		if (aNewStartAngle != aStartAngle) Change = %true ;
		}
	}
		
   if (Change) {
      	// Modify Raglan cut-out ARC
   	SetGeometry (hArc, xyAxilla,  nAxillaRadius, aNewStartAngle, aNewDeltaAngle);
   	
	// Modify Insert ARC
   	GetGeometry (hInsertArc, &xyAxilla,  &nAxillaRadius, &aStartAngle, &aDeltaAngle);
        	aNewStartAngle =  aStartAngle - aAngleChange  ;
        	aNewDeltaAngle = aDeltaAngle + aAngleChange  ;
      	SetGeometry (hInsertArc, xyAxilla,  nAxillaRadius, aNewStartAngle, aNewDeltaAngle);

	// Modify Insert ARC Seam Line
   	GetGeometry (hInsertSeamArc, &xyAxilla,  &nAxillaRadius, &aStartAngle, &aDeltaAngle);
      	SetGeometry (hInsertSeamArc, xyAxilla,  nAxillaRadius, aNewStartAngle, aNewDeltaAngle);

	// Modify Insert Line
	// Redraw as we can't SetGeometry properly on a "line" nor  can we rotate it
        	GetExtent (hInsertLine, &xyStart, &xyEnd) ;
        	xyStart = CalcXY ("relpolar", xyEnd, Calc("length",xyStart, xyEnd), Calc("angle",xyEnd, xyStart) - aAngleChange) ;
	Execute("menu", "SetLayer", Table("find", "layer", "Template" + sArm)); 
                hEnt = AddEntity("line", xyStart, xyEnd) ;
                SetDBData (hEnt,"ID", sInsertLineID);
                DeleteEntity(hInsertLine) ;
	}
	
   Execute ("menu", "ViewRedraw") ;
   Exit(%ok, "Finished") ;
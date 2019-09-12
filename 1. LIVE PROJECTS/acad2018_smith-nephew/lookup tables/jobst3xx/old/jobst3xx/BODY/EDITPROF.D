// File Name:	EDITPROF.D
// Purpose:	Edit the profile at the  axilla ht
//		Bodysuit
// Version:	1.01
// Date:		6.June.97
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action

HANDLE
	hEnt,
	hCutOutCurve,
	hCutOutMarker,
	hBackCurve,
	hBackCurveMarker,
	hChan,
	hCurve,
	hLayer;

XY
	xyCutOutMarker,
	xyBackCurveMarker,
	xyPt1,
	xyPt2,
	xyTmp,
	xyTmpBCM,
	xyTmpCOM ;

STRING
	sClass,
	sProfileID,
	sText,
	sTmp, sTmp1,
	sTmpBCM, 
	sTmpCOM,
	sSide;

NUMBER
	nStringLength,
	nFound,
	nType,
	nRequired,
	ii, 
	nn,
	nTol,
	nBraDeltaY,
	iVertexCount,
	iVertexCutOutCurve,
	nCutOutCurveDelta,
	iVertexBackCurve,
	nBackCurveDelta;

ANGLE
	aTmpBCM,
	aTmpCOM ;

BOOL
	bBraFound ;
	
// Clear user selection
//
   UserSelection ("clear") ;
   UserSelection ("update") ;
   
   STRING		sPathJOBST;
   sPathJOBST = GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;
   SetData("PathMacro",  sPathJOBST + "\\BODY" );
  
// Get profile, identify side 
//
   if (! hCurve=GetEntity ("Select Bodysuit Profile", "all")) Exit (%ok, "User Cancelled") ;

   Display ("cursor", "wait", "Searching for more data");
   GetDBData (hCurve, "ID", &sProfileID);
   nStringLength = StringLength(sProfileID);
   if (StringCompare( "Left", StringMiddle(sProfileID, nStringLength - 3, 4))){
	sSide = "Left";
	}
    if (StringCompare( "Right", StringMiddle(sProfileID, nStringLength - 4, 5))) {
	sSide = "Right";
	}

// Check if this is a valid Profile exit if not
// This is crude check only
    if (StringLength (sSide) == 0) Exit (%cancel, "A profile  not found\n") ;

// Get data for a DORSAL Zipper 
   sTmp = "DB ID = '" + sProfileID + "'  AND  ( DB curvetype = 'Bra' OR DB curvetype ='BackCurve' OR DB curvetype ='BackCurveMarker' OR DB curvetype='CutOutCurve'   OR DB curvetype = 'CutOutMarker')";
   nFound = 0;  nReqired = 4;  bBraFound = %false;
   hChan = Open ("selection", sTmp );
   if (hChan) {
	ResetSelection(hChan);
	while (hEnt = GetNextSelection(hChan)) {
		GetDBData(hEnt, "curvetype", &sTmp);
		if (StringCompare(sTmp, "BackCurve")) {
			nFound = nFound + 1 ;
			hBackCurve = hEnt ;
			}
		if (StringCompare(sTmp, "BackCurveMarker")) {
			nFound = nFound + 1 ;
			hBackCurveMarker = hEnt ;
			GetGeometry(hEnt, &sTmpBCM, &xyBackCurveMarker , &xyTmpBCM, &aTmpBCM) ;
			}
		if (StringCompare(sTmp, "CutOutCurve" )){
			nFound = nFound + 1 ;
			hCutOutCurve = hEnt ;
			}
		if (StringCompare(sTmp, "Bra" )){
			bBraFound = %true;
			UserSelection ("add", hEnt) ;
			}
		if (StringCompare(sTmp, "CutOutMarker" )){
			nFound = nFound + 1 ;
			hCutOutMarker= hEnt ;
			GetGeometry(hEnt, &sTmpCOM, &xyCutOutMarker , &xyTmpCOM, &aTmpCOM) ;
			}
		}
   	}
   Close("selection", hChan); 

// Check that sufficent data have been found, otherwise exit
//
    if (nFound < nReqired) Exit (%cancel, "Missing data for selected Bodysuit!\n") ;
    if (nFound > nReqired) Exit (%cancel, "Two or more drawings of the same style exist!\nDelete the extra drawing/s and try again.\n") ;

// Make sure both curves are polyines or curves else we will have problems later
// with GetVertexCount etc.
   GetEntityClass (hCutOutCurve, &sClass) ;
   if (!StringCompare("curve", sClass) && !StringCompare("polyline", sClass)) Exit (%cancel, "CutOut Curve is not a curve or polyline!\n") ;
   
   GetEntityClass (hBackCurve, &sClass) ;
   if (!StringCompare("curve", sClass) && !StringCompare("polyline", sClass)) Exit (%cancel, "Back Curve is not a curve or polyline!\n") ;

// We have found the curves and the markers we now check to see which curve has been displaced
// relative to it's marker. 
// NB the use of a tolerance of 0.1" in both X and Y
   nTol = 0.1 ;

   ii = 1;
   iVertexCutOutCurve = 0 ;	// Explicity 0
   nCutOutCurveDelta = 0;	// Explicity 0
   
   iVertexCount = GetVertexCount(hCutOutCurve) ;
   while ( ii <= iVertexCount) {
	GetVertex (hCutOutCurve, ii, &xyPt1) ;
	// Check this point against xyCutOutMarker 
	if (abs (xyCutOutMarker.x - xyPt1.x) <= nTol) {
		iVertexCutOutCurve = ii ;
		nCutOutCurveDelta = xyCutOutMarker.y - xyPt1.y ;
		break;
		}
	ii = ii + 1;
	}
   
   ii = 1;
   iVertexBackCurve = 0 ; 	// Explicity 0
   nBackCurveDelta = 0 ; 	// Explicity 0
  
   iVertexCount = GetVertexCount(hBackCurve) ;
   while ( ii <= iVertexCount) {
	GetVertex (hBackCurve, ii, &xyPt1) ;
	// Check this point against xyCutOutMarker 
	if (abs (xyBackCurveMarker.x - xyPt1.x) <= nTol) {
		iVertexBackCurve = ii ;
		nBackCurveDelta = xyBackCurveMarker.y - xyPt1.y ;
		break;
		}
	ii = ii + 1;
	}
    
// Having got the Curve Vertex and Displacements (Delta)
// We need to resolve any posible conflicts  before we proceed

// Editable vertex not found
   if ( iVertexCutOutCurve == 0 ) Exit (%cancel, "Can't find a vertex to move in the CutOutCurve!\nPerhaps the curve vertex is not on the axilla construction line") ;
   if ( iVertexBackCurve == 0 ) Exit (%cancel, "Can't find a vertex to move in the BackCurve!\nPerhaps the curve vertex is not on the axilla construction line") ;

// No displacement of either curve, No change
   if ((abs(nBackCurveDelta) <= nTol)  && (abs(nCutOutCurveDelta) <= nTol) ) Exit (%cancel, "Neither curve has been displaced!\nNo change to be made.") ;

// Both curves displaced, Can't resolve this 
   if ((abs(nBackCurveDelta) > nTol)  && (abs(nCutOutCurveDelta) > nTol) ) Exit (%cancel, "Both curves have been displaced!\nCan't resolve this conflict.") ;

// Modify as required
   if ((abs(nBackCurveDelta) >= nTol) ) {
   	// BackCurve Changed
	// Therefor change CutOut Curve and Marker
	GetVertex (hCutOutCurve,iVertexCutOutCurve , &xyPt1) ;
   	xyPt1.x = xyCutOutMarker.x ;
   	xyPt1.y = xyPt1.y + (nBackCurveDelta / 2) ;
	SetVertex (hCutOutCurve,iVertexCutOutCurve , xyPt1) ; 	
	SetGeometry(hCutOutMarker, sTmpCOM, xyPt1 , xyTmpCOM, aTmpCOM) ;
	
	// Reset Back Curve marker
	GetVertex (hBackCurve,iVertexBackCurve , xyPt2) ; 	
	SetGeometry(hBackCurveMarker, sTmpBCM, xyPt2 , xyTmpBCM, aTmpBCM) ;
	
	// Delta Y for bras (if required)
	nBraDeltaY = nBackCurveDelta / 2 ;
   	}
   else {
   	// CutOutCurve Changed
	// Therefor change Back Curve and Marker
	GetVertex (hBackCurve, iVertexBackCurve , &xyPt1) ;
   	xyPt1.y = xyPt1.y + (nCutOutCurveDelta/2) ;
	SetVertex (hBackCurve, iVertexBackCurve , xyPt1) ; 	
	SetGeometry(hBackCurveMarker, sTmpBCM, xyPt1, xyTmpBCM, aTmpBCM) ;
	
	// Reset CutOutCurve marker
	GetVertex (hCutOutCurve, iVertexCutOutCurve , xyPt2) ; 	
	SetGeometry(hCutOutMarker, sTmpCOM, xyPt2 , xyTmpCOM, aTmpCOM) ;
	
	// Delta Y for bras (if required)
	nBraDeltaY = -nCutOutCurveDelta ;
	}

// if a Bra has been found then move the Bra Selection set
//     
   if ( bBraFound) Transform( "move", "translate", 0, nBraDeltaY );
   	
// Reset and exit
//
   UserSelection ("clear") ;
   UserSelection ("update") ;
   Execute ("menu", "SetLayer",  Table("find", "layer", "1")) ;
   Exit (%ok, "Bodysuit Profile Edit Complete");

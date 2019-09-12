// File Name:	GLVTHUMB.D
// Purpose:	Transfer thumb curve on palm to thumb
//		piece
//		Manual  Glove
// Version:	1.01
// Date:		14.May.1996
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action

HANDLE
	hDlg,
	hEnt,
	hTPCurve,
	hTPSeam,
	hNewTPSeam,
	hTPOEdge,
	hNewTPOEdge,
	hChan,
	hCurve,
	hLayer;

XY
	xyPt1,
	xyPt2,
	xyPt3,
	xyPt4,
	xyStart,
	xyEnd,
	xyTmp,
	xyTmp1,
	xyTmp2;

STRING
	sClass,
	sText,
	sUnits,
	sProfileID,
	sTmp, sTmp1,
	sSide;

NUMBER
	nFound,
	nRequired,
	nOffSet,
	ii, nn;
ANGLE
	aAngle1, 
	aAngle2;
	

// Clear user selection
//
   UserSelection ("clear") ;
   UserSelection ("update") ;
   
   STRING		sPathJOBST;
   sPathJOBST = GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;
   SetData("PathMacro",  sPathJOBST + "\\GLOVEMAN" );
   
// Get profile, identify side 
//
   if (! hCurve=GetEntity ("Select Glove Profile", "all")) Exit (%ok, "User Cancelled") ;

   Display ("cursor", "wait", "Searching for more data");
   GetDBData (hCurve, "ID", &sProfileID);
   nStringLength = StringLength(sProfileID);
   if (StringCompare( "Left", StringMiddle(sProfileID, nStringLength - 3, 4))){
	sSide = "Left";
	nOffSet = 0.125;
	}
    if (StringCompare( "Right", StringMiddle(sProfileID, nStringLength - 4, 5))) {
	sSide = "Right";
	nOffSet = -0.125;
	}

// Check if this is a valid Glove exit if not
// This is crude check only
    if (StringLength (sSide) == 0) Exit (%cancel, "A Glove Profile was not selected\n") ;

// Get data for Thumb 
   sTmp = "DB ID = '" + sProfileID + "' AND  (DB curvetype = 'ThumbPalmCurve' OR DB curvetype = 'ThumbPieceSeam' OR DB curvetype = 'ThumbPieceOutsideEdge' )";
   nFound = 0 ; nReqired = 3;
   hChan = Open ("selection", sTmp );
   if (hChan) {
	ResetSelection(hChan);
	while (hEnt = GetNextSelection(hChan)) {
		GetDBData(hEnt,"curvetype", &sTmp) ;
		if (StringCompare(sTmp, "ThumbPalmCurve")) {
			nFound = nFound + 1 ;
			hTPCurve = hEnt ;
   			}
		if (StringCompare(sTmp, "ThumbPieceSeam")) {
			nFound = nFound + 1 ;
			hTPSeam = hEnt ;
			}
		if (StringCompare(sTmp, "ThumbPieceOutsideEdge" )){
			nFound = nFound + 1 ;
			hTPOEdge = hEnt ;
			}
		}
   	}
   Close("selection", hChan); 

// Check that sufficent data have been found, otherwise exit
//
    if (nFound < nReqired) Exit (%cancel, "Missing data for selected Glove!\n") ;
    if (nFound > nReqired) Exit (%cancel, "Two or more drawings of the same style exist!\nDelete the extra drawing/s and try again.\n") ;

// fail if not enough vertex in the curve, minimum of 4 required.
//
   nVertex = GetVertexCount (hTPCurve) ;
   if (nVertex <= 4) Exit (%cancel, "Thumb Curve on Palm does not contain enough points to complete action!");
   

// DEFINE Functions and Subroutines
//
// Function:  	BOOLEAN = FN_LinLinInt( xyLineStart1, xyLineEnd1, xyLineStart2, xyLineEnd2);
//
// Parameters:	xyLineStart1 = %1, %2
//		xyLineEnd1 = %3, %4
//		xyLineStart2 = %5, %6
//		xyLineEnd2 = %7, %8
//
// Returns:	%true if intersection found
//		%false if no intesection
//		xyInt =  intersection ( Global Variable )
//
		 
// Variables:
// ------------------------------------------ Start -------------------------------------------
XY	
	xyInt; 	// if it exist, intersection of line and line
NUMBER 
	nK1, nK2, nM1, nM2, nSlope1, nSlope2, nCase, 
	nY, nX;
// ------------------------------------------ End   -------------------------------------------

function FN_LinLinInt()
{
  // Caclculate slope
   nCase = 0 ;
   nSlope1 = Calc("angle", %1, %2, %3, %4);
   if ( nSlope1 ==  0 || nSlope1 == 180 ) nCase = nCase + 1 ;
   if ( nSlope1 ==90 || nSlope1 == 270 ) nCase = nCase + 2 ;

   nSlope2 = Calc("angle", %5, %6, %7, %8);
   if ( nSlope2 == 0  || nSlope2 == 180 ) nCase = nCase + 4 ;
   if ( nSlope2 ==90 || nSlope2 == 270 ) nCase = nCase + 8 ;

  // Parallel orthogonal lines, no intersection possible
   if ( nCase == 5 || nCase == 10) return (%false); 
   
  // Line1 is Horizontal and the Line2 is Vertical 
   if ( nCase == 9 ) {
   	nX = %5;
   	nY = %2;
	}
  // Line1 is Vertical and the Line2 is Horizontal 
   if ( nCase == 6 ) {
   	nX = %1;
   	nY = %6;
	}
	
  // Line 1 is Horizontal or Line 2 is horizontal 
     if (nCase == 1 || nCase == 4 ) {
      	if (nCase == 1) {
   		nM1 = (%8 - %6) / (%7 - %5) ;		// Slope
   		nK1 = %6 - nM1*%5;		// Y-Axis intercept
		nY = %2;
		}
	else{
   		nM1 = (%4 - %2) / (%3 - %1) ;		// Slope
   		nK1 = %2 - nM1*%1;		// Y-Axis intercept
		nY= %6;
		}
   		
   	// Solve for X at the given Y value
      	nX = (nY - nK1) / nM1;
	}
	
  // Line 1 is Vertical or Line 2 is Vertical 
	
     if (nCase == 2 || nCase == 8 ) {
      	if (nCase == 2) {
   		nM1 = (%8 - %6) / (%7 - %5) ;		// Slope
   		nK1 = %6 - nM1*%5;		// Y-Axis intercept
		nX = %1;
		}
	else{
   		nM1 = (%4 - %2) / (%3 - %1) ;		// Slope
   		nK1 = %2 - nM1*%1;		// Y-Axis intercept
		nX= %5;
		}
   		
   	// Solve for Y at the given X value
 	nY = (nM1*nX) + nK1;
	}

  // Both lines are Non-Orthogonal Lines
   if ( nCase==0  ) {
   	nM1 = (%4 - %2) / (%3 - %1) ;		// Slope
   	nM2 = (%8 - %6) / (%7 - %5) ;		// Slope
   	if (nM1 == nM2) return (%false) ;
   	nK1 = %2 - nM1*%1;		// Y-Axis intercept
   	nK2 = %6 - nM2*%5;		// Y-Axis intercept
   	if (nK1 == nK2) return (%false) ;
	// Find X
	nX = (nK2 - nK1) / (nM1 - nM2);	
	// Find Y				
 	nY = (nM1*nX) + nK1;
 	}
   
   xyInt.x = nX;
   xyInt.y = nY;
   
   return (%true);
   
   } // End function FN_LinLinInt

 
// Get XY translation
//
   GetVertex(hTPCurve, 1, &xyPt1) ;
   GetVertex(hTPCurve, GetVertexCount (hTPCurve), &xyPt2) ;
   if (xyPt1.y < xyPt2.y) 
  	xyStart = xyPt1;
   else
  	xyStart = xyPt2;
   GetVertex(hTPSeam, 1, &xyPt1) ;
   GetVertex(hTPSeam, GetVertexCount (hTPSeam), &xyPt2) ;
   if (xyPt1.y < xyPt2.y) 
   	xyEnd = xyPt1;
   else
   	xyEnd = xyPt2;

//
//
   UserSelection ("clear") ;
   UserSelection ("update") ;
   UserSelection ("add", hTPCurve) ;

// Translate the curve 
//
   Transform ("copy", "translate", xyEnd.x - xyStart.x , xyEnd.y - xyStart.y) ;

// Get the entity handle of the last created
// Set it's Data Base Values
   hNewTPSeam = UID ("find", UID ("getmax")) ;
   SetDBData(hNewTPSeam, "curvetype" , "ThumbPieceSeam");
   SetDBData(hNewTPSeam, "ID" , sProfileID);
   
	
// Loop through the Vertext in the ThumbPieceSeam offsetting them by the seam allowance
// nOffSet is above, it's value depending on the side. 
   GetVertex ( hNewTPSeam, 1, &xyTmp) ; 
   GetVertex ( hNewTPSeam, 2, &xyTmp1) ; 
   GetVertex ( hNewTPSeam, 3, &xyTmp2) ; 
   
   StartPoly("fitted") ;
    GetVertex ( hTPOEdge, 1, &xyInt) ;
    AddVertex (xyInt) ; 
   
    nn = 3;    
    while (nn <= nVertex) {
   	// Offset Outside edge from Seam
   	aAngle1 = Calc ("angle", xyTmp, xyTmp1) + 90;
   	aAngle2 = Calc ("angle", xyTmp1, xyTmp2) + 90 ;
   	xyPt1= CalcXY ("relpolar", xyTmp, nOffSet ,aAngle1);
   	xyPt2= CalcXY ("relpolar", xyTmp1, nOffSet ,aAngle1);
   	xyPt3= CalcXY ("relpolar", xyTmp1, nOffSet ,aAngle2);
   	xyPt4= CalcXY ("relpolar", xyTmp2, nOffSet ,aAngle2);
   	if ( FN_LinLinInt (xyPt1, xyPt2, xyPt3, xyPt4)) {
  		AddVertex (xyInt) ; 
		}
   	nn = nn + 1;
   	xyTmp = xyTmp1 ;
   	xyTmp1 = xyTmp2 ;
   	GetVertex ( hNewTPSeam, nn, &xyTmp2) ; 
  	}
  	
    GetVertex ( hTPOEdge, GetVertexCount (hTPOEdge), &xyInt) ;
    AddVertex (xyInt) ; 
   EndPoly() ;
   
// Find the new Outside Edge and set it's values
//   
   hNewTPOEdge = UID ("find", UID ("getmax")) ;
   SetDBData(hNewTPOEdge, "curvetype" , "ThumbPieceOutsideEdge");
   SetDBData(hNewTPOEdge, "ID" , sProfileID);
   SetEntityData (hNewTPOEdge, "layer", Table("find", "layer", "Template" + sSide));
   
// Clean up
//            
   DeleteEntity(hTPSeam) ;
   DeleteEntity(hTPOEdge) ;
   UserSelection ("clear") ;
   UserSelection ("update") ;

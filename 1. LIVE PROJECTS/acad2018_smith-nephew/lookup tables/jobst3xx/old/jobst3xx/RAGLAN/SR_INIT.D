// File Name:	VR_INIT.D
// Purpose:	Add Vest Raglan to Arm
//		
// Note:		Not a user called routine.
//		Called from the Arm drawing macro created by Visual Basic
//
// Version:	1.01
// Date:		1.Nov.94
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date	By	Action
// Input
//

// Initialise raglan drawing
//
//N.B. If hSleeveProfile is not given then the control points xyAxilla and xyLowerAxilla
//	are given explicity.
//	This allows this code to be reused by the side neck scoop tool
   if (hSleeveProfile) {
	// Establish proximal control points from Sleeve profile
   	GetEntityClass( hSleeveProfile, &sEntClass) ;
  	if (StringCompare("curve", sEntClass) || StringCompare("polyline", sEntClass)) {
   		nTape = GetVertexCount (hSleeveProfile) ;
		GetVertex( hSleeveProfile, nTape, &xyAxilla) ;
		}
   	else
		{
		GetGeometry(hSleeveProfile, &xyProfileStart, &xyAxilla) ;
		}

	xyLowerAxilla.y = xyOrigin.y ;
   	xyLowerAxilla.x = xyAxilla.x ;
   	}
   	

// Get 1/3 rd line down from upper proximal scale and intersect with 
//
   nLen = (Calc ("length", xyLowerAxilla, xyAxilla) - nSeam) / 3 ;
   
   if (FN_CirLinInt (	xyAxilla.x          , xyAxilla.y - nLen,
		xyAxilla.x + 100, xyAxilla.y - nLen,
		xyAxilla,
		nAxillaBackNeckRad))
	{	
	xyBackNeck = xyInt;
	xyEndBottomCurve.y = xyBackNeck.y - nShoulderToBackRaglan;
 	xyEndBottomCurve.x = xyBackNeck.x ;
	}
   else 	
	Exit (%cancel, "Can't form raglan curve with this sleeve data\n") ;
	
// set layer

    if (StringCompare(sSleeve, "Right"))
	PRSetLayerTemplateRight() ;
    else
	PRSetLayerTemplateLeft() ;

// Get vest curve angle

// Calculate the required vestcurve angle.
// Open the VESTCURVE data file
// Check for 1. Existance and 2. Corruption  
//

   hCurve = Open ("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly");
   if (hCurve) {
	if (GetLine(hCurve, &sLine)){
		ScanLine( sLine, "blank", &nLength, &aAngle);
		}
	else
		{
		Exit (%abort, "Can't read VESTCURV.DAT\nFile maybe corrupted");
		}
	}
   else{
	Exit (%abort, "Can't open VESTCURV.DAT\nCheck installation");
	}

// Intersect the unrotated vest curve with a radius, that is centered on the 
// 0,0, to get a curve angle.
// This angle will be used to calculate vestcurve rotation.
// For both the upper and lower curves
//

   nLowerCurveRadius = Calc ("length", xyLowerAxilla, xyEndBottomCurve) ;
   aUpperCurve = 100000 ;	// Impossible value used to test for non-intersecton
   aLowerCurve = 100000 ;
   aPrevAngle = 0;
   xyPt1.x = 0;
   xyPt1.y = 0;
   while (GetLine(hCurve,&sLine)) {
	aAngle = aAngle + aPrevAngle;
	xyPt2 = CalcXY("relpolar", xyPt1, nLength, aAngle);
	if ( FN_CirLinInt( xyPt1, xyPt2, 0.0, 0.0, nAxillaBackNeckRad ))
		aUpperCurve = Calc ("angle", 0,0, xyInt);  	
	if ( FN_CirLinInt( xyPt1, xyPt2, 0.0, 0.0, nLowerCurveRadius ))
		aLowerCurve = Calc ("angle", 0,0, xyInt);  	
	xyPt1 = xyPt2;
	aPrevAngle = aAngle;
	ScanLine(sLine, "blank", &nLength, &aAngle);	
  	}
   if (aUpperCurve == 100000 || aLowerCurve == 100000) {
	Display("message", "error","Sleeve drawing error\nCan't make Ragalan curve with this data");
	}
   Close ("file", hCurve);

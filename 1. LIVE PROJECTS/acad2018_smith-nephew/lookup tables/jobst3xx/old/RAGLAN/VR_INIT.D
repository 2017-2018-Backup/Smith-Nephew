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
//            	sAxillaType, either 	Regular 2"
//                                 	Regular 1½"
//                                   	Regular 2½"
//                                  	Open
//                                  	Mesh
//                                 	Lining
//            	nAxillaFrontNeckRad
//            	nAxillaBackNeckRad
//            	nShoulderToBackRaglan
//	hSleeveProfile
//	sSleeve
//	xyOrigin		
//	sFileNo

HANDLE
	hChan,
	hEnt,
	hBody,
	hCurve,
	hMesh;
STRING
	sEntClass,
	sTmp;

NUMBER
	nLowerCurveRadius;

STRING
	sAxillaType;

NUMBER
	ii, jj,
	nLoop,			//
	nInt,
	nDec,
	nOk,
  	nError,
	nValue,
	nPreValue,
	nRadius,
 	DrawSegment,
	DrawInsert,
	nUnitsFac;			// The conversion factor based on
					// the inputs units type

// Variables for Axilla drawing macros
XY
	xyProfileStart,
	xyLowerAxilla,			// Axilla on CL of profile template
//	xyAxilla,
//	xyAxillaLow,			// Lowest Axilla point on profile
	xyAxillaConstruct_1,
	xyAxillaConstruct_2,
	xyInsertConstruct_1,
	xyInsertConstruct_2,
	xyInsertConstruct_3,
	xyInsertConstruct_4,
	xyBackNeck,
	xyRaglanNeck,
	xyRaglanAxilla,
	xyEndBottomCurve,
	xyStartBottomCurve,
	xyTangent,
	xyMeshAtSeamAllowance,
	xyCen;

NUMBER
	nSeam,		// Seam allownce
	nSeamAllowance,
	nLen;

XY
	xyInt;		// Global, returned intersection point from FN_CirLinInt()

STRING
	sLine,
	sError,
	sMeshSymbol,
	sCurvFileNo;

XY
	xyPt1,
	xyPt2,
	xyPt,
	xyTmp,
	xyCen;

ANGLE
	aAngle,
	aPrevAngle,
	aLowerCurve,
	aUpperCurve;
 
// Load Defaults
//
    nRegAxillaNormal = 2.0;
    nRegAxillaChild = 1.5;
    nRegAxillaLargeAdult = 2.5;
    nSeamAllowance = 0.125;
    nSeam = 0.1875 ;
    nRegAxillaFac_1 = 1.5 ;			// GOP 01-02/18, 16.1
    nRegAxillaFac_2 = 0.75;			// GOP 01-02/18, 16.3
    nFrontNeckRedFac = 0.125 ;			// GOP 01-02/18, 15.5
    nShouldertoBackRagRedFac = 0.1875;	// GOP 01-02/18, 15.6

    nMeshAxillaGussetFac = 0.75 ;
    nMSLengthGusset = 2.1875;
    nMeshAxillaGussetBoysFac =  0.8125 ;
    nMSLengthGussetBoys = 2.375;


// Layer setting procedures
//

procedure PRSetLayerTemplateRight () {
   hLayer = Table("find", "layer", "TemplateRight") ;
   if ( hLayer > %zero )
 	Execute ("menu", "SetLayer",  hLayer) ;
   else
   	Table ("add", "layer", sLayer, "Paper Template Right Hand side", "current", Table("find", "color", "Red"));
}

procedure PRSetLayerTemplateLeft () {
   hLayer = Table("find", "layer", "TemplateLeft") ;
   if ( hLayer > %zero )
 	Execute ("menu", "SetLayer",  hLayer) ;
   else
   	Table ("add", "layer", sLayer, "Paper Template Left Hand side", "current", Table("find", "color", "Blue"));
}

procedure PRSetLayerData () {
   hLayer = Table("find", "layer", "Data") ;
   if ( hLayer > %zero )
 	Execute ("menu", "SetLayer",  hLayer) ;
   else
   	Table ("add", "layer", sLayer, "Data Layer", "current", Table("find", "color", "Black"));
}

procedure PRSetLayerNotes () {
   hLayer = Table("find", "layer", "Notes") ;
   if ( hLayer > %zero )
 	Execute ("menu", "SetLayer",  hLayer) ;
   else
   	Table ("add", "layer", sLayer, "Notes in Green", "current", Table("find", "color", "Green"));
}


procedure PRSetLayer1 () {
   hLayer = Table("find", "layer", "1") ;
   if ( hLayer > %zero ) Execute ("menu", "SetLayer",  hLayer) ;
}

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
		 



// Procedure PRDataStamp ( Origin.x, Origin.y, nStampType)
// Purpose:	To add a stamp containing the patient details 
//		at the given origin
//
// Arguments:	%1 = Origin.x
//		%2 = Origin.y
//		%3 = nStampType where :-
//			21 = Sleeve, Axilla insert left abbreviated stamp
//			22 = Sleeve, Axilla insert right abbreviated stamp
//			
// Notes:		The origin is the top of the text
//		text is printed downwards
	
	
procedure PRDataStamp( ) {
   SetData("TextFont", 0);
   SetData("TextVertJust", 32);		// Bottom
   SetData("TextHorzJust", 1);		// Left
   SetData("TextHeight", 0.1);
   SetData("TextAspect", 0.6);

// Axilla Inserts stamp SLEEVE
//
    if ( %3 == 21 || %3 == 22) {
   	if ( %3 == 21)    AddEntity("text", "Sleeve Left" + "\n" + sName + "\n" + sWorkOrder  ,%1 ,%2);
		else  AddEntity("text", "Sleeve Right" + "\n" + sName + "\n" + sWorkOrder ,%1, %2);
    }

} // End PRDataStamp() 



// Function:
//	nRounded Value = FNRound ( nLength )
// Returns:
//	nRounded Value	- The result rounding. 
//			  Based on the "measurement" options set in DRAFIX

// Arguments:
//	nLength

function FNRound () {
	return ( Value("length", Format("length", %1)));
} /* End FNRound */



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

// File Name:	WH_LEG1.D
// Purpose:	Draw the Waist High with Legs
//		Draws 1st leg and foot
//		Draws Patient data
//
// Version:	1.01
// Date:		05.Feb.94
// Author:	Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date	By	Action
// Oct.95	gg	Revised w.r.t CAD to Triton link
//
   Display ("cursor", "wait", "Drawing");

HANDLE
	hChan,
	hTitle,			// TITLEBOX
	hWHBody,		// WAISTBOX
	hLeg,			// LEGBOX
	hTemplateLayer,
	hCurv,
	hEnt,
	hLayer;

NUMBER
	nFirstTape,		
	nLastTape,
	nAnkleTape;


// TITLEBOX Details
STRING
	sPatient,			// Patient name from title symbol box
 	sDiagnosis,		// Diagnosis        "      "         "        "
	sSEX,			// Patient SEX      "     "          "       "
	sFileNo,			// File Number     "      "         "        "
	sUnits,			// The type of units
	sAge;			// Patients age     "     "         "        "
				//     the inputs units type
XY
	xyTitleOrigin, 
	xyTitleScale;
ANGLE
	aTitleAngle;

// WHBODYBOX  Details
STRING
	sTOSCir,			// Top Of Support Circumference
	sLeftThighCir,		// Proximal Thigh Left @ fold of buttocks
	sRightThighCir,		// Proximal Thigh Right @ fold of buttocks
	sThighGivenRed,
	sFoldHt,			// Height to fold of buttocks
	sFootLength,
	sCrotchStyle,		// Crotch Style 1 of 7
	sToeStyle,
	sLegStyle,
	sFabric;			// Fabric Type combination of 
				//     material and modulus

NUMBER
	nOpenBack,		//From WHBODDIA via DB field Body
	nOpenFront,		//  	"	"	"
	nCrotchFrontFactor,		//  	"	"	"
	nUseRevisedHtsFlag,	//  	"	"	"
	nTOSHtRevised,		//  	"	"	"
	nWaistHtRevised,		//  	"	"	"
	nFoldHtRevised,		//  	"	"	"

	nLegStyle,		//From WHBODDIA via DB field ID
	nLeftLegStyle,		//  	"	"	"
	nRightLegStyle,		//  	"	"	"

	nTOSCir,			// Numerical equivelents of above	
	nLeftThighCir,		
	nRightThighCir,		
	nFoldHt,		
 	nOtherAnkleMTapeLen,	
	nGivenRedFac;

// LEG Variables
// LEGBOX Details
// 
STRING
	sLeg,			// Left or Right leg (NB. Now Declared in shell)
	sTapeLengths,		// Joined tape lenghts
	sGramsAnkle,		// Grams calculated at ankle
	sReductionAnkle,		// Reductions calculated at ankle
	sAnkleTape,		// Designated ankle tape
	sFirstTape,		// First tape
	sLastTape,		// Last tape
	sFootPleat1,		// Pleated tape towards Foot
	sTopLegPleat1,		// Pleated tape towards Top of Leg
	sFootPleat2,		// Pleated tape towards Foot
	sTopLegPleat2,		// Pleated tape towards Top of Leg
	sModulus,		// Modulus of Choosen fabric
	sMaterial,		// Material
	sMMAnkle,		// Chosen MM based on diagnosis
	sPressure,		// Chosen Template Pressure
	sReduction,		// Calculated reduction etc at each leg tape		
	sStretch,			// fron WHFIGURE.EXE
	sTapeMMs,		// ditto
	sFootLength,
	sToeStyle,
	sFootLabel;

NUMBER
	nElastic,		// Elasitic flag for panty 0=BelowKnee, 1=AboveKnee
	nFirstTape,		
	nLastTape,
	nAnkleTape,
	nHeelLength,
	nAnkleLength,
	nFootLength;

// Main Control points
//
XY
	xyCO_WaistBott,		// For Label of a BRIEF at the waist
	xyO,
	xyOBody,		// The origin of the Body 
	xyTOS,			// Top Of Support
	xyWaist,			// Waist
	xyMidPoint,		// MidPoint or Cleft
	xyLargest,		// Largest part of buttocks
	xyLastTape,
	xyFoldPanty,
	xyThigh,
	xyFold;			// Fold or Groin Height
	
// Foot Points (Variables)
XY
	xyFirstTape,
	xyHeel,
	xyAnkle,
	xyAnkleM,
	xyAnkleMPrev,
	xyHeelCntrDistal,
	xyHeelCntrMidDistal,
	xyHeelCntrMidProximal,
	xyHeelCntrProximal,
	xyToeCntrLow,
	xyToeCntrMid,
	xyToeCntrHigh,
	xyToePoint,
	xyToePnt,		// Extreme Left point
	xyToeOFF,		// Straight toes
	xyToeSeam,
	xyToeStop,
	xyToeCL;

// Factors, Tollerances and References
//
NUMBER
	nSeam;			// Standard seam allowance

STRING
	sTextList,		// Tape text in a string
   	sSpecificationNO,		// Specification used to design WH
   	sSpecificationDate;		// Date of above specification


// Working variables
//

STRING
	sCrotchText,
	sGussetSize,
	sLayer,			// Tmp type variable
	sTmp,			//    "      "         "
	sLine,			//    "      "         "
	sError,			// Error reporting string
	sConversion,
	sRef,
	sScale,
	sFile,
	sID,
	sSymbol,
	sThing,
	sText ;

NUMBER
	ii, jj, nn,
	nNo,
	nSpace,
	n20Len,
	nReduction,
	nLoop,			
	nInt,
	nDec,
	nOk,
  	nError,
	nValue,
	nPreValue;

XY
	xyOtemplate,
	xyStart,
	xyEnd,
	xyPt1,
	xyPt2,
	xyInt,
	xyTmp;

ANGLE
	aAngle,
	aPrevAngle;
// FLAGS
BOOL
	Loop,
	PantyLeg,
	Male,
	LeftLeg,
	RightLeg,
	FootLess,
	Female,
	BodyFound,
	ClosedCrotch,
	OpenCrotch,
	BigAnkle;		//*
STRING
	sWorkOrder,
	sPathJOBST,
	sFirstLeg,
	sSecondLeg;

// Clear User any user selections
//
   UserSelection ("clear") ;
   Execute ("menu", "SetStyle",  Table("find", "style", "bylayer")) ;
   Execute ( "menu", "SetStyle", Table( "find",  "style", "bylayer"));


//





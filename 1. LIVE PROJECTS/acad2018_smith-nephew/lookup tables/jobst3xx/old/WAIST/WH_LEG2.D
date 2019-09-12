// File Name:	WH_LEG2.D
// Purpose:	Draw the Waist High with Legs 2nd leg
//		Declares variables etc
//
// Version:	1.01
// Date:		07.Feb.94
// Author:	Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date	By	Action

   Display ("cursor", "wait", "Drawing 2nd Leg");

HANDLE
	hChan,
	hTitle,			// TITLEBOX
	hWHBody,		// WAISTBOX
	hLeg,			// LEGBOX
	hTemplateLayer,
	hFirstLegCurve,
	hFileCurve,
	hCurv,
	hEnt,
	hLayer;

NUMBER
	nFirstTape,		// Numerical equivelants of above
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

NUMBER
	nAge,			// Numerical version of Pateints Age
	nUnitsFac;		// The conversion factor based on
				//     the inputs units type
// WHBODYBOX  Details
STRING
	sLeftThighCir,		// Proximal Thigh Left @ fold of buttocks
	sRightThighCir,		// Proximal Thigh Right @ fold of buttocks
	sTOSCir,
	sCrotchStyle,
	sFoldHt;
				//     material and modulus
NUMBER
	nOpenOff,		//From CO_CenterArrow - Marker
	nOpenBack,		//From WHBODDIA via DB field Body
	nOpenFront,		//  	"	"	"
	nCrotchFrontFactor,		//  	"	"	"
	nUseRevisedHtsFlag,	//  	"	"	"
	nTOSHtRevised,		//  	"	"	"
	nWaistHtRevised,		//  	"	"	"
	nFoldHtRevised,		//  	"	"	"

	nLegStyle,		//From WHBODDIA via DB field ID
	nLeftLegStyle,		//  	"	"	"
	nRightLegStyle;		//  	"	"	"


// LEG Variables
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
	sFabric,			// Fabric Type combination of 
				//     material and modulus
	sReduction, 
              	sTapeMMs,
            	sStretch,
	sFootLength,
	sLegStyle,
	sToeStyle;

NUMBER
	nElastic,		  	// Elasitic flag for Panty 0=BelowKnee, 1=AboveKnee
	nFirstTape,		
	nLastTape,
	nAnkleTape,
	nHeelLength,
	nAnkleLength;

// Main Control points
//
XY
	xyFilletTop,
	xyFilletBott,
	xyO,			// The origin of the Body 
	xyNull,
	xyTOS,			// Top Of Support
	xyWaist,			// Waist
	xyMidPoint,		// MidPoint or Cleft
	xyLargest,		// Largest part of buttocks
	xyButtockArcCen,
	xyCO_LargestTop,
	xyCO_LargestBott,
	xyCO_TOSBott,
	xyCO_TOSTop,
	xyCO_WaistBott,
	xyCO_WaistTop,
	xyCO_MidPointBott,
	xyCO_MidPointTop,
	xyCO_ArcCen,
	xyCO_OpenArcTop,
	xyCO_OpenArcBott,
	xyCO_CenterArrow,
	xyLastTape,
	xyFoldPanty,
	xyThigh,
	xyFold;			// Fold or Groin Height
	
ANGLE
	aCO_ArcStart,
	aCO_ArcStartOriginal,
	aCO_ArcDelta,
	aCO_OpenArcTopStart,
	aCO_OpenArcTopDelta,
	aCO_OpenArcBottStart,
	aCO_OpenArcBottDelta;

NUMBER
	nFilletRadius,
 	nOtherAnkleMTapeLen,	
	nTopArcToLargestOffset,
	nBottArcToLargestOffset,
	nCO_ArcRadius,
	nCO_OpenArcTopRadius,
	nCO_OpenArcBottRadius;


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
	// Factors loaded by WHLG2DEF.D
	// 
	nFootLength,
	nSeam;			// Standard seam allowance

STRING
	sTextList,		// Tape text in a string
   	sSpecificationNO,		// Specification used to design WH
   	sSpecificationDate,		// Date of above specification
	sExtTemplt1TolRef,		// Body Extension template positioning
	sExtTemplt2TolRef,		//    "	    "	   "	    "
	sExtTemplt3TolRef;		//    "	    "	   "	    "


// Working variables
//

STRING
	sLayer,			// Tmp type variable
	sTmp,			//    "      "         "
	sLine,			//    "      "         "
	sError,			// Error reporting string
	sRef,
	sScale,
	sThing,
	sType,
	sFile,
	sID,
	sSymbol,
	sFootLabel,
	sFirstLeg, sSecondLeg,
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
	nX, nY,
  	nError,
	nValue,
	nPreValue;

XY
	xyOFirstBody,
	xyOBody,
	xyLargestFirstBody,
	xyOtemplate,
	xyStart,
	xyEnd,
	xyPt1,
	xyPt2,
	xyInt,
	xyTmp;

ANGLE
	aAngle,
	aTransAngle,
	aStartAngle,
	aPrevAngle;
// FLAGS
BOOL
	AddPoint,
	Loop,
	PantyLeg,
	BodyFound,
	Male,
	LeftLeg,
	RightLeg,
	Female,
	FootLess,
	OpenTopArcFound,
   	OpenBottArcFound,
	BigAnkle;		//*
	
STRING
	sWorkOrder,
	sPathJOBST;

// Clear User any user selections
//
   UserSelection ("clear") ;
   Execute ("menu", "SetStyle",  Table("find", "style", "bylayer")) ;
   Execute ( "menu", "SetStyle", Table( "find",  "style", "bylayer"));


// End of WH_LEG2.D

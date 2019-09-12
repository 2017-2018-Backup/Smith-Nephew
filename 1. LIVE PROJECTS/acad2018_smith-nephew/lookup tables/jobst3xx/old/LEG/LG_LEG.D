// File Name:	LG_LEG.D
// Purpose:	Draw Leg (Loads Variables)
//
// Version:	1.01
// Date:		09.Dec.94
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date	By	Action

   Display ("cursor", "wait", "Drawing");

HANDLE
	hChan,
	hTemplateLayer,
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
	sAge,			// Patients age     "     "         "        "
	sTitleName;

NUMBER
	nAge,			// Numerical version of Pateints Age
	nUnitsFac;		// The conversion factor based on
				//     the inputs units type
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
	sFabric,			// Fabric Type combination of 
				//     material and modulus
	sFootLength,
	sLegStyle,
	sHeelStyle,
	sToeStyle;

NUMBER
	nFirstTape,		// Numerical equivelants of above
	nLastTape,
	nAnkleTape,
	nHeelLength,
	nAnkleLength;


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
	nFootLength,
	nSeam;			// Standard seam allowance

STRING
	sTextList,		// Tape text in a string
   	sSpecificationNO,		// Specification used to design WH
   	sSpecificationDate;		// Date of above specification


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
	sSymbol,
	sFootLabel,
	sFirstLeg, sSecondLeg,
	sText,
	sWorkOrder;

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
 	nOtherAnkleMTapeLen,	// Dummy to allow use of WHFTPNT.D
	nPreValue;

XY
	xyO,
	xyProfileLast,
	xyThighPlt,
	xyPrevProfileLast,
	xyProfileStart,
	xyOtemplate,
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
	Male,
	LeftLeg,
	RightLeg,
	Female,
	FootLess,
	BigAnkle;		//*

// Clear User any user selections
//
   UserSelection ("clear") ;
   UserSelection ("update") ;
   Execute ("menu", "SetStyle",  Table("find", "style", "bylayer")) ;
   Execute ("menu", "SetStyle",  Table("find", "color", "bylayer")) ;





// File Name:	WH_W4OC.D
// Purpose:	Draw the Waist High One Leg Open 
//		Crotch, ie a W4OC
// Version:	1.01
// Date:		19.Dec.94
// Author:	Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date	By	Action

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
	nAnkleTape,		// From WHLEGDIA.D via DB field AnkleTape
	nAnkleMM, 		//	"	"	"	"
	nAnkleGrams,		//	"	"	"	"
	nAnkleReduction,		//	"	"	"	"
	nHeelLength,		//	"	"	"	"
	nAnkleLength;		//	"	"	"	"

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
ANGLE
	aTitleAngle;

XY 	
	xyTitleScale,
	xyTitleOrigin;

// WHBODYBOX  Details
STRING
	sTOSCir,			// Top Of Support Circumference
	sTOSGivenRed,		// User Supplied Reduction
	sTOSHt,			// Top Of Support Height
	sWaistCir,		// Waist Circumference
	sWaistGivenRed,		// User Supplied Reduction
	sWaistHt,		// Waist Height
	sMidPointCir,		// Midpoint Circumference
	sMidPointGivenRed,	// User Supplied Reduction
	sMidPointHt,		// Midpoint Height
	sLargestCir,		// Largest part of buttocks Circumference
	sLargestGivenRed,		// User Supplied Reduction
	sLargestHt,		// Largest part of buttocks Height
	sLeftThighCir,		// Proximal Thigh Left @ fold of buttocks
	sRightThighCir,		// Proximal Thigh Right @ fold of buttocks
	sThighGivenRed,
	sFoldHt,			// Height to fold of buttocks
	sCrotchStyle,		// Crotch Style 1 of 7
	sLegStyle;

NUMBER
	nTOSCir,			// Numerical equivelents of above	
	nTOSHt,			
	nWaistCir,		
	nWaistHt,		
	nMidPointCir,	
	nMidPointHt,	
	nLargestCir,	
	nLargestHt,		
	nLeftThighCir,		
	nRightThighCir,		
	nFoldHt,	
	nGivenRedFac;

// LEG Variables
// LEGBOX Details
// 
STRING
	sLeg,			// Left or Right leg 
	sGramsAnkle,		// From WHLEGDIA.D via DB field AnkleTape
	sReductionAnkle,		// 	"	"	"	"
	sAnkleTape,		//  	"	"	"	"
	sFirstTape,		// First tape
	sLastTape;		// Last tape

NUMBER
	nFirstTape,		// Numerical equivelants of above
	nLastTape;
	
// Main Control points
//
XY
	xyFilletTop,
	xyFilletBott,
	xyO,			// The origin of the Body 
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
	xyFold,			// Fold or Groin Height
	xyCO_CutMark,
	xyCutMark;
	
ANGLE
	aTopMark,
	aCO_ArcStart,
	aCO_ArcStartOriginal,
	aCO_ArcDelta,
	aCO_OpenArcTopStart,
	aCO_OpenArcTopDelta,
	aCO_OpenArcBottStart,
	aCO_OpenArcBottDelta;

NUMBER
	nFilletRadius,
	nTopArcToLargestOffset,
	nBottArcToLargestOffset,
	nCO_ArcRadius,
	nCO_OpenArcTopRadius,
	nCO_OpenArcBottRadius;

// Factors, Tollerances and References
//
NUMBER
	// Factors loaded by WHCUTDEF.D
	// 
	nSeam,			// Standard seam allowance
	nOpenOff,
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

	nStandardRedFac,		// 0.83,
	nBabyRedFac,		// 0.88, 1 to 3 yrs of age
	nChildRedFac,		// 0.86, 4 to 10 yrs of age
	nUpliftRedFac,		// 0.88
	nReductionTol,		// 0.05 Tolerance +/- (Max. Allowable) 
	nFemaleOpenCrotchFac,	// Cut out positioning, Open Crotch (Front)
	nMaleOpenCrotchFac,	//    "    " 	       "	        "	"
	nChildTo3OpenCrotchFac,	//    "    " 	       "	        "	"
	nFemaleClosedCrotchFac,	// Cut out positioning, Closed Crotch (Front)
	nMaleClosedCrotchFac,	//    "    " 	       "	        "	" 
	nChildTo3ClosedCrotchFac,	//    "    " 	       "	        "	" 

	// Template Extension positioning "Case" values
	// Dependant on nFoldHt.
	nExtTempltTol1max,	// Case 1
	nExtTemplt1tape,		// 	
	nExtTempltTol2min,	// Case 2
	nExtTempltTol2max,	//
	nExtTemplt2tape,		//
	nExtTempltTol3min,	// Case 3
	nExtTemplt3tape,		//

	nBodyFrontStep,	 	// Step that the following will be applied by  
   	nBodyFrontIncreaseOff,	// Amount to increase for each step smaller
   	nBodyFrontReduceOff,	// Amount to decrease for each step larger

	nCutOutConstructOff_1,	// Cut Out Construction offset 1
   	nCutOutDiaMaxTol,		// Maximum Cut Out Size
  	nCutOutConstructFac_1,	// Cut Out Construction Factor 1

	nBodyBackCutOutMinTol,	// Back Body Cut Out Minimum
   	nEndBackBodyOff,		// Back of body extension

	nPantyThighRed1,
	nPantyThighRed2,


    	// Calculated Factors from input data
	//
	nTOSBackOff,
	nCutOutDiameter,
	nReductionFac,		// Dependant on age
	nFoldFiguredHt ;		// Dependant on nFoldHt


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
	sFile,
	sSymbol,
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
	xyO,
	xyOtemplate,
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
	IgnoreMidPoint_CO,
	Loop,
	PantyLeg,
	Male,
	LeftLeg,
	RightLeg,
	Female,
	ClosedCrotch,
	OpenCrotch;

STRING
	sWorkOrder,
	sPathJOBST;
	
// Clear User any user selections
//
   UserSelection ("clear") ;
   Execute ("menu", "SetStyle",  Table("find", "style", "bylayer")) ;


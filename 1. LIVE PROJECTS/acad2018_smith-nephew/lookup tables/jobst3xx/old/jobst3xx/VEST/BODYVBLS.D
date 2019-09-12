// File Name:	BODYVBLS.D
// Purpose:	Declarations of variables for BODY.D
//		to reduce the BODY.D file to managable size
//	
// Version:	1.01
// Date:		18.Jan.94
// Author:		Gary George

// Measurements for all styles
NUMBER
	nLtSCir,			// Left Shoulder Circumference
	nRtSCir,			// Right Shoulder Circumference
	nSCir,				// Shoulder Circumference
	nNeckCir,			// Neck Circumference	
	nSWidth,			// Shoulder Width
	nS_Waist,		// Shoulder to Waist
	nChestCir,		// Chest Circumference
	nWaistCir,		// Waist Circumference
	nS_EOS,			// Shoulder to End Of Support 	(Optional)
	nEOSCir;			// End Of Support Circumference (Optional)

// BRA details. 
NUMBER
	nS_Breast,		// Shoulder to Just Under Breast	
	nBreastCir,		// Circumference Just Under Breast
	nNippleCir;		// Circumference at Nipple line
STRING
	sBraLtCup,		// Left Bra Cup size
	sBraRtCup,		// Right Bra Cup size
	sBraLtDisk,		// Bra disk for left cup
	sBraRtDisk;		// Bra disk for right cup
// AXILLA Details.
NUMBER
	nAxilla;			// Flag to indicate that Left and Right
				// should be figured seperatly
				// 2 = Both; where 1=Left and 2=Right.
				// 1 = Single Axilla
STRING
	sID,
	sLtAxillaType,		// Left Axilla type 
	sRtAxillaType,		// Right Axilla type
	sAxillaType,		// Used when looping through Axilla
	sAxillaTypeList;		// List of Axilla types supported eg. Mesh, Open etc


// NECK details. 
NUMBER
	nNeckActualCir,		// Actual calculated value
	nNeckGiven,		// Given neck circumference
	nNeckDimension,
	nTurtleLength,
	nNeckCir,			// Neck Circumference
	nCLNeckDrop;		// Revised neck based on 16.Oct.97 modifications
STRING
	sSide,
	sNeckType,		// Neck Type eg. Regular, Scoop, Turtle etc.
	sBackNeckType,
	sNeckDimension,
	sBackNeckDim,
	sNeckNotes;		// Neck notes to be written on template

//CLOSURE 
STRING
	sClosure;			// Body Closure 

// PATIENT Details
NUMBER
	nAge;			// Patient age
STRING
	sPatient,			// Patient name
	sSEX,			// Patients SEX
	sWorkOrder,
	sDiagnosis,		// Patent diagnosis
	sAge,			// Patients year of birth
	sFileNo;			// Paients file number

//CONVERSIONS / FIGURING
NUMBER
	nUnitsFac,		// Units conversion factor
	nEOSCirFac,		// End Of Support conversion factor
	nWaistCirFac,		// Waist Circumference conversion factor
	nBreastCirFac,		// Just under breast circumference conversion factor
	nChestCirFac,		// Chest Circumference conversion factor
	nSCirFac,			// Shoulder to axilla conversion factor
	nNeckFrontFac,		// Offset from top of shoulder line to raglan intersection
	nNeckFac_1,		// Neck Circumference conversion factor (1 of 3)
	nNeckFac_2,		// Divide by number (2 of 3)
	nNeckFac_3,		// Neck seam allowance (3 of 3)
	nNeckFac_4,		// Scoop neck factor
	nNeckFac_5,		
	nBreastCirUserFac,		// User suplled breast circumference factor
	nEOSCirUserFac,		// User suplled EOS circumference factor
	nWaistCirUserFac,		// User suplled Waist circumference factor
	nRegAxillaFac_1,		// Regular axilla radius
	nRegAxillaFac_2,		// Regular axilla construction radius
	nRegAxillaNormal,		// ditto -  adult
	nRegAxillaLargeAdult,	// ditto -  large adult
	nRegAxillaChild,		// ditto -  child
	nHighSLineFac,		// Addition factor to produce highest shoulder line
	nBraFac_1,		// left offset from  axilla shoulder line
	nBraMinProfile, 		// minimum Bra Cup to profile
	nBraMinEOS,		// minimum to EOS to bottom of Bra Cup
	nBraPosFacX,		// Top of bra disk 1, used to offset scale position for 
	nBraPosFacY;		//   the other disks

STRING
	sUnitsType,		// Units type
	sFabric;

//General Variables
XY
	xyO,
	xyPt1,			// Tmp variables
	xyPt2,			//    "         "
	xyCen,			//    "         "
	xyBackNeckCL, 		// Back neck on Centre Line
	xyWaistCL,
	xyWaistOFF,
	xyEOSCL,
	xyEOSOFF,
	xyFrontNeckCL,
	xyFrontNeckOFF,
	xyFrontNeckOthers,
	xyRaglanNeckOriginal,
	xyFrontNeckCen,
	xyBackNeckConstruct,	// Back neck Construction point
	xyBackNeckConstruct_2,	// Back neck Construction point
	xyBackNeckCen,		// Center of back neck radius
	xyRaglanNeck,
	xyRaglanAxilla,
	xyHighestAxilla,
	xyLowestAxilla,
	xyAxilla,
	xyAxillaLow,
	xyBreast,
	xySleeveLess,
	xyInsertConstruct_1,
	xyInsertConstruct_2,
	xyInsertConstruct_3,
	xyInsertConstruct_4,
	xyAxillaConstruct_2,
	xyAxillaLowRight,
	xyMeshAtSeamAllowance,
	xyNeck_HighS_Int;
	
HANDLE
	hOriginalLayer,
	hCurve,			// VESTCURV data file 
	hMesh;			// Mesh Axilla symbol

NUMBER
	nLength,			// Tmp variable
	nRadius,			//   "        "
	nCount,			//   "        "
	DrawSegment,		// Flag to turn VESTCURVE drawing on
	DrawInsert,		// Flag to turn Axilla INSERT drawing on and off
	nLowSLine,		// Lowest point of shoulder line
	nHighSLine,		// Highest point of shouder line
	nBackNeckRadius,		// Radius of the back neck line
	nAxillaFrontNeckRad,	// Radius from Axilla to Front Neck (Used to calculate aVestAngle)
	nAxillaBackNeckRad,	// Radius from Axilla to Back Neck (passed to sleeve)
	nShoulderToRaglan,	// Top shoulder line to back raglan length (passed to sleeve)
	nAxillaFrontNeckRadRight,	// As above for right sleeve
	nAxillaBackNeckRadRight,	// As above for right sleeve
	nShoulderToRaglanRight,	// As above for right sleeve
				// (GOP 01-02/18, 15.4)
	nDistanceAlongRaglan,
	EOSGiven,		// Flag to indicate that an EOS measurement has been given
	BRAGiven,		// Flag to indicat that Bra measurements have been given
	nSeamAllowance,		// General Seam allowance.
	nBraAxillaHt,		// Height of top of bracup
	nBraCLOffset,		// Offset from template CL
	nBraScale,		// Bra Disk scaling factor relative to disk 
	nDiskXoff,		// X & Y Position of top of bra disk relative
	nDiskYoff,		// to the symbol insert position
	nInt,
	nInitial_nAxilla,
	nDec;

ANGLE
	aCurveRotation,		// Variable from FN_CurveAngle  
	aAngle,			// Tmp variable
	aPrevAngle,		//    "        "
	aCurve,			//    "        "
	aVestCurve,		// Vest Curve angle, Forms raglan
	aHighestCurveRotation,
	aLowestCurveRotation,
	aHighestVestCurve,
	aLowestVestCurve;

STRING
	sHighestAxilla,
	sLtHandColor,		// Template Left hand side colour
	sRtHandColor,		//        "        Right         "
	sMeshSymbol,
	sTmp,
	sError,
	sLine;			// Tmp variable

BOOL
	SleeveLess,
	bPrevAxillaWasMesh;
	
HANDLE
	hEnt ;
	
// BODYBRA Variables
XY	xyDisk;
STRING  	sDisk;

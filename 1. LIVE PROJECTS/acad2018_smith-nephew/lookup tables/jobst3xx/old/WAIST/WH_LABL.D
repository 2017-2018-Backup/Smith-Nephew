// File Name:	WH_LABL.D
// Purpose:	Labels 1st leg with crotch details
//
// Version:	1.01
// Date:		14.Mar.94
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
	hEnt,
	hLayer;

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
	nRightThighCir,
	nLeftThighCir,
	nAge,			// Numerical version of Pateints Age
	nUnitsFac;		// The conversion factor based on
				//     the inputs units type
XY
	xyTitleOrigin, 
	xyTitleScale;
ANGLE
	aTitleAngle;



// WHBODYBOX  Details
STRING
	sCrotchStyle,
	sLegStyle,
	sLeftThighCir,		// Proximal Thigh Left @ fold of buttocks
	sRightThighCir;		// Proximal Thigh Right @ fold of buttocks

NUMBER
	nLeftThighCir,		
	nRightThighCir,	
	nLegStyle,		//From WHBODDIA via DB field ID
	nLeftLegStyle,		//  	"	"	"
	nRightLegStyle,		//  	"	"	"
	nMaxCir;			//  	"	"	"


// Main Control points
//
XY
	xyCO_CenterArrow,
	xyCO_LargestTop,
	xyCO_LargestBott;


// Factors, Tollerances and References
//
NUMBER
	// Factors loaded by WHLG1DEF.D
	// 
	nSeam;			// Standard seam allowance
	

STRING
	sLeg,
	sCO_CenterArrowID,
	sCO_LargestTopID,
 	sCO_LargestBottID,
   	sSpecificationNO,
   	sSpecificationDate,
	sCrotchText,
	sGussetSize,
	sLayer,			// Tmp type variable
	sTmp,			//    "      "         "
	sType,
	sLine,			//    "      "         "
	sError,			// Error reporting string
	sRef,
	sScale,
	sFile,
	sSymbol,
	sText;

NUMBER
	ii, jj, nn,
	nNo,
	nSpace,
	n20Len,
	nCO_ArcDiameter, 
	nLoop,			
	nInt,
	nDec,
	nOk,
  	nError,
	nValue,
	nPreValue;

XY
	xyPt1,
	xyPt2,
	xyTmp;

ANGLE
	aAngle,
	aPrevAngle;
// FLAGS
BOOL
	Loop,
	PantyLeg,
	Male,
	Female,
	LeftLeg,
	LeftCO_CenterArrowFound,
	RightLeg,
	ClosedCrotch,
	OpenCrotch;

// Find JOBST installed directory
// Set path to macros
//
   STRING	 sPathJOBST;
   sPathJOBST = GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;
   SetData("PathMacro",  sPathJOBST + "\\WAIST" );
   
// Clear User any user selections
//
   UserSelection ("clear") ;
   Execute ("menu", "SetStyle",  Table("find", "style", "bylayer")) ;

// Load Functions and Procedures
// Load Factors, Tolerances and References
//
   @WHLBLDEF.D;

// Get titlebox details Note - Reuse of sleeve based function  
//
   hChan = Open ("selection", "DB SymbolName = 'mainpatientdetails'");
   ResetSelection(hChan);
   hTitle = GetNextSelection(hChan);
   if (hTitle) {
	GetDBData( hTitle, "patient", &sPatient);
   	GetDBData( hTitle, "fileno", &sFileNo);
 	GetDBData( hTitle, "age",  &sAge);
		nAge = Value("scalar", sAge);
   	GetDBData( hTitle, "sex", &sSEX);
	GetDBData( hTitle, "units", &sUnits);
	// Set up units conversion factor
   	if ( StringCompare(sUnits, "cm")) nUnitsFac = 10/25.4;	// Cm to Inches   
   	if ( StringCompare(sUnits, "inches")) nUnitsFac = 1;	// Inches

   	GetDBData( hTitle, "diagnosis", &sDiagnosis);
	GetGeometry(hTitle, &sTitleName, &xyTitleOrigin, &xyTitleScale, &aTitleAngle);
  	Close("selection",hChan);

  	}
        else {
	Exit(%cancel, "Patient details cannot be found\nPlease ensure that TITLEBOX has been used\nThen try again\n");
	}
   Close("selection", hChan);

// Setup SEX flags
   Male = %false;
   Female = %false;
   if (StringCompare("Male", sSEX)) Male = %true;
	else Female = %true;

// Get body details from "waistbody"
//
   @WHLBLBOD.D; 	

// Get details from the markers
//
   @WHLBLMKR.D;

// Add Crotch Label
//
   @WHLBLCRO.D;

  Exit (%ok, "Label Complete");




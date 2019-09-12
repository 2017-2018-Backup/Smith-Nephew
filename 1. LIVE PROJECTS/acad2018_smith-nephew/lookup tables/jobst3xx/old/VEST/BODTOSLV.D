// File Name:	BODTOSLV.D
// Purpose:	Transfer from the drawn vest details of the 
//		Axilla to back neck
//		Axilla to front neck
//		Back neck to Back neck construction point
//
// Version:	1.01
// Date:		21.June.94
//
// Author:	Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action

HANDLE
	hChan,
	hTitle,
	hLayer,
	hBody,
	hEnt;

// TITLEBOX Details
STRING
	sPatient,			// Patient name from title symbol box
 	sDiagnosis,		// Diagnosis        "      "         "        "
	sSEX,			// Patient SEX      "     "          "       "
	sFileNo,			// File Number     "      "         "        "
	sUnits,			// The type of units
	sAge;			// Patients age     "     "         "        "

// Body box details
NUMBER
	ii, jj,
	nAxillaFrontNeckRad,	// Radius from Axilla to Front Neck 
	nAxillaBackNeckRad,	// Radius from Axilla to Back Neck 
	nLowerCurveRadius,
	nShoulderToRaglan;	// Top shoulder line to back raglan length 

STRING
	sTmp,
	sLayer,
	sSleeve,
	sLtAxillaType,		// Left Axilla type 
	sRtAxillaType;		// Right Axilla type

XY
	xyAxilla,
	xyFrontNeck,
	xyBackNeck,
	xyBackNeckConstruct;

//
   UserSelection ("clear") ;
   UserSelection ("update") ;

// get line / arcs etc to establish side

   hEnt = GetEntity ("Select Vest Raglan Profile", "line arc curve polyline");
   if(!hEnt)Exit(%ok,"Vest Raglan Profile not selected");

   GetEntityData ( hEnt, "layer" , &hLayer); 
   GetTableData ( "layer", hLayer, "name", &sLayer);
   sSleeve = StringMiddle(sLayer, 9, (StringLength(sLayer) - 8));

   if (!StringCompare(sSleeve,"Left") && !StringCompare(sSleeve, "Right")) 
	Exit (%cancel, "Select a Right or Left Raglan profile only\n\n") ;

   if( !GetUser ("xy", "Axilla Point", &xyAxilla))Exit(%ok,"");
   if( !GetUser ("xy", "Front Neck and Raglan intesection", &xyFrontNeck))Exit(%ok,"");
   if( !GetUser ("xy", "Back Neck at end of Raglan", &xyBackNeck))Exit(%ok,"");
   if( !GetUser ("xy", "Back Neck at Highest Shoulder line", &xyBackNeckConstruct))Exit(%ok,"");

   nAxillaFrontNeckRad = Calc ("length", xyAxilla, xyFrontNeck);
   nAxillaBackNeckRad =  Calc ("length", xyAxilla, xyBackNeck);
   nShoulderToBackRaglan =  Calc ("length", xyBackNeck, xyBackNeckConstruct);

// Set bodybox details
//
   sTmp =  "DB SymbolName = 'vestbody'" ; 
   hChan = Open ("selection", sTmp);
   ResetSelection(hChan);
   hBody = GetNextSelection(hChan);
   if (hBody) {
   	if (StringCompare(sSleeve, "Right")) {
	   	sTmp = MakeString ("scalar",nAxillaFrontNeckRad);
		  SetDBData ( hBody,  "AFNRadRight", sTmp) ;
		sTmp = MakeString ("scalar",nAxillaBackNeckRad ) ;
		  SetDBData ( hBody,  "ABNRadRight", sTmp) ;
		sTmp = MakeString ("scalar", nShoulderToBackRaglan ) ;
		  SetDBData ( hBody,  "SBRaglanRight", sTmp) ;
		}
	else {
	   	sTmp = MakeString ("scalar",nAxillaFrontNeckRad);
		  SetDBData ( hBody,  "AxillaFrontNeckRad", sTmp) ;
 		sTmp = MakeString ("scalar",nAxillaBackNeckRad ) ;
   		  SetDBData ( hBody,  "AxillaBackNeckRad", sTmp) ;
 		sTmp = MakeString ("scalar", nShoulderToBackRaglan ) ;
   		  SetDBData ( hBody,  "ShoulderToBackRaglan", sTmp) ;
		}
	}
        else {
        	Close ("selection", hChan);
    	Exit(%cancel, "Body data cannot be found\nPlease ensure that BODY details have been given\nThen try again\n");
    	}

Close ("selection", hChan);

//
// Get the required points

 Exit (%ok, "Data Transfer Finished");




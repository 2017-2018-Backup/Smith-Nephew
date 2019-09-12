// File Name:	WH_DRAW1.D <sDrawType>
// Purpose:	General drawing tool
//		Pokes all Data to a VB program
//		VB program then starts the drawing
//
// Version:	1.01
// Date:	01.Dec.94
// Author:	Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
// 18.Oct.95	GG	Revised to use arguments 
// 27.Sep.96	GG	Use while loop 

   Display ("cursor", "wait", "Loading");

// Arguments
// sDrawType can be: - 
//
//	"WHCUT"		Cutout for and body
//	"WH1LG"		1st Leg drawing
//	"WH2LG"		2nd Leg drawing

HANDLE
	hDDE,
	hChan,
	hTitle,			// TITLEBOX
	hWHBody,		// WHBODYBOX
	hSym,
	hLeg;			// LEGBOX 

STRING
	sPathJOBST,
	sFileNo,
	sLeg,
	sTmp,
	sSymSelection,
	sDrawType,
	sSymName,
	sTmp1;
XY
	xySymO,
	xyScale;
ANGLE
	aAngle;

// Find JOBST installed directory
//
   sPathJOBST =  GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;

// Clear User any user selections
//
   UserSelection ("clear") ;
  
   sDrawType = GetArgument (1) ;

// Get titlebox details 
//
   hChan = Open ("selection", "DB SymbolName = 'mainpatientdetails'");
   ResetSelection(hChan);
   hTitle = GetNextSelection(hChan);
   if (!hTitle) {
	Close("selection",hChan);
	Exit(%cancel, "Patient details cannot be found\nPlease ensure that a TITLEBOX has been given\nThen try again\n");
	}
   else
   	GetDBData( hTitle, "fileno", &sFileNo);

   Close("selection", hChan);

   sSymSelection = "(DB SymbolName = 'waistleg'  AND DB fileno = '" + sFileNo + "') OR " +
   	 	"(DB SymbolName = 'waistbody'  AND DB fileno = '" + sFileNo + "')" ;

// Start the application and get DDE link
//
   Execute ("application",sPathJOBST + "\\waist\\whdraw", "normal");
   hDDE = Open ("dde", "whdraw", "whdraw");

// Get LEFT & RIGHT legbox handle and waistbody symbol
//
   hChan = Open ("selection",sSymSelection );
   if ( hChan && hDDE) {
          ResetSelection(hChan);
          while (hSym = GetNextSelection(hChan)) {
             	GetGeometry(hSym, &sSymName, &xySymO, &xyScale, &aAngle);
      	if(StringCompare(sSymName, "waistleg" )){ 
		hLeg = hSym ;
      		GetDBData( hLeg, "leg", &sLeg);
      		
  		GetDBData( hLeg, "TapeLengthsPt1", &sTmp); 
  		GetDBData( hLeg, "TapeLengthsPt2", &sTmp1); 
			sTmp = sTmp + sTmp1 ;
		Poke ( hDDE, "txt" + sLeg + "Lengths", sTmp );

  		GetDBData( hLeg, "TapeMMs", &sTmp); 
  		GetDBData( hLeg, "TapeMMs2", &sTmp1); 
			sTmp = sTmp + sTmp1 ;
		Poke ( hDDE, "txt" + sLeg + "MMs", sTmp );

    		GetDBData( hLeg, "Reduction", &sTmp); 
  		GetDBData( hLeg, "Reduction2", &sTmp1); 
			sTmp = sTmp + sTmp1 ;
		Poke ( hDDE, "txt" + sLeg + "Red", sTmp );

  		GetDBData( hLeg, "Grams", &sTmp); 
  		GetDBData( hLeg, "Grams2", &sTmp1); 
			sTmp = sTmp + sTmp1 ;
		Poke ( hDDE, "txt" + sLeg + "Str", sTmp );

		GetDBData(hLeg, "AnkleTape", &sTmp);
		Poke ( hDDE, "txt" + sLeg + "Ankle", sTmp);

  		GetDBData( hLeg, "Pressure", &sTmp);
		Poke ( hDDE, "txt" + sLeg + "Template", sTmp);

   		GetDBData( hLeg, "FootPleat1", &sTmp);
		Poke ( hDDE, "txt" + sLeg + "FootPleat1", sTmp);

   		GetDBData( hLeg, "FootPleat2", &sTmp);
		Poke ( hDDE, "txt" + sLeg + "FootPleat2", sTmp);

   		GetDBData( hLeg, "TopLegPleat1", &sTmp);
		Poke ( hDDE, "txt" + sLeg + "TopLegPleat1", sTmp);

   		GetDBData( hLeg, "TopLegPleat2", &sTmp);
		Poke ( hDDE, "txt" + sLeg + "TopLegPleat2", sTmp);

   		GetDBData( hLeg, "ToeStyle", &sTmp);
		Poke ( hDDE, "txt" + sLeg + "ToeStyle", sTmp);

   		GetDBData( hLeg, "FootLength", &sTmp);
		Poke ( hDDE, "txt" + sLeg + "FootLength", sTmp);

  		GetDBData( hLeg, "Data", &sTmp);
		Poke ( hDDE, "txt" + sLeg + "Data", sTmp);

  		GetDBData( hLeg, "Fabric", &sTmp); 
		Poke ( hDDE, "txt" + sLeg + "Fabric", sTmp);

		Poke ( hDDE, "txtUid" + sLeg + "Leg", MakeString("long",UID("get",hLeg)));
		}

      	if(StringCompare(sSymName, "waistbody" )){ 
      		hWHBody = hSym;
  		GetDBData( hWHBody, "Fabric", &sTmp); 
		Poke ( hDDE, "txtFabric", sTmp);

   		GetDBData( hWHBody, "TOSCir", &sTmp); 
		Poke ( hDDE, "txtTOSCir", sTmp);

   		GetDBData( hWHBody, "TOSGivenRed", &sTmp); 
		Poke ( hDDE, "txtTOSRed", sTmp);

   		GetDBData( hWHBody, "TOSHt", &sTmp); 
		Poke ( hDDE, "txtTOSHt", sTmp);

   		GetDBData( hWHBody, "WaistCir", &sTmp);
		Poke ( hDDE, "txtWaistCir", sTmp);

   		GetDBData( hWHBody, "WaistGivenRed", &sTmp);
		Poke ( hDDE, "txtWaistRed", sTmp);

   		GetDBData( hWHBody, "WaistHt", &sTmp);
		Poke ( hDDE, "txtWaistHt", sTmp);

   		GetDBData( hWHBody, "MidPointCir", &sTmp);
		Poke ( hDDE, "txtMidPointCir", sTmp);

   		GetDBData( hWHBody, "MidPointGivenRed", &sTmp);
		Poke ( hDDE, "txtMidPointRed", sTmp);

   		GetDBData( hWHBody, "MidPointHt", &sTmp);
		Poke ( hDDE, "txtMidPointHt", sTmp);

   		GetDBData( hWHBody, "LargestCir", &sTmp);
		Poke ( hDDE, "txtLargestCir", sTmp);

   		GetDBData( hWHBody, "LargestGivenRed", &sTmp);
		Poke ( hDDE, "txtLargestRed", sTmp);

   		GetDBData( hWHBody, "LargestHt", &sTmp);
		Poke ( hDDE, "txtLargestHt", sTmp);

   		GetDBData( hWHBody, "LeftThighCir", &sTmp);
		Poke ( hDDE, "txtLeftThighCir", sTmp);

   		GetDBData( hWHBody, "RightThighCir", &sTmp);
		Poke ( hDDE, "txtRightThighCir", sTmp);

   		GetDBData( hWHBody, "ThighGivenRed", &sTmp);
		Poke ( hDDE, "txtThighRed", sTmp);

   		GetDBData( hWHBody, "FoldHt", &sTmp);
		Poke ( hDDE, "txtFoldHt", sTmp);

   		GetDBData( hWHBody, "CrotchStyle", &sTmp);
		Poke ( hDDE, "txtCrotchStyle", sTmp);

   		GetDBData( hWHBody, "ID", &sTmp);
		Poke ( hDDE, "txtLegStyle", sTmp);

   		GetDBData( hWHBody, "Body", &sTmp);
		Poke ( hDDE, "txtBody", sTmp);

		Poke ( hDDE, "txtUidBody", MakeString("long",UID("get",hWHBody)));
 	}
            }
    }

  Close("selection", hChan);

    if (hTitle && hDDE) {
	Poke ( hDDE, "txtDrawType", sDrawType);

 	GetDBData( hTitle, "age",  &sTmp);
	Poke ( hDDE, "txtAge", sTmp);

 	GetDBData( hTitle, "fileno",  &sTmp);
	Poke ( hDDE, "txtFileNo", sTmp);

	GetDBData( hTitle, "units", &sTmp);
	Poke ( hDDE, "txtUnits",sTmp);

	GetDBData( hTitle, "Patient", &sTmp);
	Poke ( hDDE, "txtPatientName",sTmp);

	GetDBData( hTitle, "Diagnosis", &sTmp);
	Poke ( hDDE, "txtDiagnosis",sTmp);

	GetDBData( hTitle, "Age", &sTmp);
	Poke ( hDDE, "txtAge",sTmp);

	GetDBData( hTitle, "SEX", &sTmp);
	Poke ( hDDE, "txtSex",sTmp);
	
	GetDBData( hTitle, "WorkOrder", &sTmp);
	Poke ( hDDE, "txtWorkOrder",sTmp);

	Poke ( hDDE, "txtUidTitle", MakeString("long",UID("get",hTitle)));
  	}

    Close ( "dde", hDDE);
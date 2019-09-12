// File Name:	WH_INBOD.D
// Purpose:	Get Data from FAX for the  Waist High
//		Based on Form 506-Lower Extremity
//		Start Dialogue Boxes for BODY.
//		
//
// Version:	2.01
// Date:		07.Jan.94
// Author:		Gary George
//
//----------------------------------------------------------
// REVISIONS:	
// Date		By	Action
//----------------------------------------------------------
// 8.Aug.94	GG	Revised to use VB Input routine
// 20.Nov.94	GG	DRAFIX 3.0 Modifications
//

   Display ("cursor", "wait", "Drawing");

HANDLE
	hDDE,
	hChan,
	hTitle,			// TITLEBOX
	hWHBody;		// WHBODYBOX

STRING
	sPathJOBST,
	sFileNo,
	sTmp;
	
// Find JOBST installed directory
//
   sPathJOBST =  GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;

// Clear User any user selections
//
   UserSelection ("clear") ;

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

// Find Waist Ht symbol 
//
   hChan = Open ("selection", "DB SymbolName = 'waistbody' AND DB fileno = '" + sFileNo + "'");
   hWHBody = %zero ;
   if (hChan) {
	ResetSelection(hChan);
   	hWHBody = GetNextSelection(hChan) ;
	}
   Close("selection", hChan); 

// Start the application and get DDE link
//
    Execute ("application", sPathJOBST  + "\\waist\\whboddia", "normal");
    hDDE = Open ("dde", "whboddia", "whboddia");

// Get and Poke waist ht details
//
   if (hWHBody && hDDE) {
   
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

	Poke ( hDDE, "txtUidWHBody", MakeString("long",UID("get",hWHBody)));
	}

   if (hTitle && hDDE) {
	// N.B. FileNo is from above
	Poke ( hDDE, "txtFileNo", sFileNo);

	GetDBData( hTitle, "patient", &sTmp);
	Poke ( hDDE, "txtPatientName", sTmp);
 	GetDBData( hTitle, "age",  &sTmp);
	Poke ( hDDE, "txtAge", sTmp);
	GetDBData( hTitle, "units", &sTmp);
	Poke ( hDDE, "txtUnits",sTmp);
	GetDBData( hTitle, "sex", &sTmp);
	Poke ( hDDE, "txtSex",sTmp);
	GetDBData( hTitle, "Diagnosis", &sTmp);
	Poke ( hDDE, "txtDiagnosis",sTmp);
	
	Poke ( hDDE, "txtUidTitle", MakeString("long",UID("get",hTitle)));
  	}

   Close ( "dde", hDDE);

   Exit (%ok, "Waist Height Body details transfered");




// File Name:	LG_INLEG.D <sLeg>
// Purpose:	Get Data from FAX for the  Waist High
//		Based on Form 506-Lower Extremity
//		Start Dialogue Boxes for LEG.
//		
//
// Version:	1.01
// Date:		31.Oct.94
// Author:	Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action

   Display ("cursor", "wait", "Loading");

HANDLE
	hDDE,
	hEnt,
	hChan,
	hLegTitle,
	hTitle,			// TITLEBOX
	hLeg;			// LEGBOX


STRING
	sLeg,
	sFileNo,
	sTmp,
	sTapeLengths;
	
  sLeg = GetArgument(1);
  
// Find JOBST installed directory
// Set path to macros
//
   STRING	 sPathJOBST;
   sPathJOBST = GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;
   SetData("PathMacro",  sPathJOBST + "\\LEG" ); 

// Clear User any user selections
//
    UserSelection ("clear");
    Execute ( "menu", "SetStyle", Table( "find",  "style", "bylayer"));
    Execute ( "menu", "SetColor", Table( "find",  "color", "bylayer"));

// Get titlebox details
//
   hChan = Open ("selection", "DB SymbolName = 'mainpatientdetails'");
   ResetSelection(hChan);
   hTitle = GetNextSelection(hChan);
   if (!hTitle) {
	Close("selection",hChan);
	Exit(%cancel, "Patient details cannot be found\nPlease ensure that TITLEBOX has been used\nThen try again\n");
	}
   else
   	GetDBData( hTitle, "fileno", &sFileNo);

   Close("selection", hChan);

   hLegTitle = %zero ;
   sTmp = "DB SymbolName = 'legcommon' AND DB fileno = '" + sFileNo + "'"  ;
   hChan = Open ("selection", sTmp);
   if (hChan) {
	ResetSelection(hChan);
   	hLegTitle = GetNextSelection(hChan) ;
	Close("selection", hChan);
	}
   Close("selection", hChan);

// Get legbox handle (if it exists)
//
   hLeg = %zero ;
   sTmp = "DB SymbolName = 'legleg' AND DB fileno = '" + sFileNo + "' AND  DB Leg = '" + sLeg + "'" ;
   hChan = Open ("selection", sTmp);
   if (hChan) {
	ResetSelection(hChan);
   	hLeg = GetNextSelection(hChan) ;
	Close("selection", hChan);
	}

// Start the application and get DDE link
//
    Execute ("application", sPathJOBST + "\\leg\\lglegdia", "normal");
    hDDE = Open ("dde", "lglegdia", "lglegdia");

   if (hLeg && hDDE) {
  	GetDBData( hLeg, "TapeLengthsPt1", &sTapeLengths);
  	GetDBData( hLeg, "TapeLengthsPt2", &sTmp);
		sTapeLengths = sTapeLengths + sTmp ;
	Poke ( hDDE, "txtLeftLengths", sTapeLengths );

  	GetDBData( hLeg, "AnkleTape", &sTmp);
	Poke ( hDDE, "txtLeftAnkle", sTmp);

   	GetDBData( hLeg, "ToeStyle", &sTmp);
	Poke ( hDDE, "txtToeStyle", sTmp);

   	GetDBData( hLeg, "FootPleat1", &sTmp);
	Poke ( hDDE, "txtFootPleat1", sTmp);

   	GetDBData( hLeg, "FootPleat2", &sTmp);
	Poke ( hDDE, "txtFootPleat2", sTmp);

   	GetDBData( hLeg, "TopLegPleat1", &sTmp);
	Poke ( hDDE, "txtTopLegPleat1", sTmp);

   	GetDBData( hLeg, "TopLegPleat2", &sTmp);
	Poke ( hDDE, "txtTopLegPleat2", sTmp);

   	GetDBData( hLeg, "FootLength", &sTmp);
	Poke ( hDDE, "txtFootLength", sTmp);

   	GetDBData( hLeg, "Pressure", &sTmp);
	Poke ( hDDE, "txtLeftTemplate", sTmp);

   	GetDBData( hLeg, "Fabric", &sTmp);
	Poke ( hDDE, "txtFabric", sTmp);

   	GetDBData( hLeg, "Anklet", &sTmp);
	Poke ( hDDE, "txtAnklet", sTmp);

   	GetDBData( hLeg, "ThighLength", &sTmp);
	Poke ( hDDE, "txtThighLength", sTmp);

   	GetDBData( hLeg, "KneeLength", &sTmp);
	Poke ( hDDE, "txtKneeLength", sTmp);

   	GetDBData( hLeg, "ThighBand", &sTmp);
	Poke ( hDDE, "txtThighBandAK", sTmp);

    	GetDBData( hLeg, "ThighBandBK", &sTmp);
	Poke ( hDDE, "txtThighBandBK", sTmp);

  	GetDBData( hLeg, "KneeBand", &sTmp);
	Poke ( hDDE, "txtKneeBand", sTmp);

	Poke ( hDDE, "txtChosenStyle", "-1");

	Poke ( hDDE, "txtUidLeftLeg", MakeString("long",UID("get",hLeg)));
 	}


    if (hTitle && hDDE) {
	Poke ( hDDE, "txtLeg", sLeg);

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

   if( hLegTitle  && hDDE) {

   	GetDBData( hLegTitle, "Fabric", &sTmp);
	Poke ( hDDE, "txtLegTitleFabric", sTmp);

	Poke ( hDDE, "txtUidLegTitle", MakeString("long",UID("get",hLegTitle)));
	}


    Close ( "dde", hDDE);

// File Name:	WH_INLEG.D <sLeg>
// Purpose:	Get Data from FAX for the  Waist High
//		Based on Form 506-Lower Extremity
//		Start Dialogue Boxes for LEG.
//		
//
// Version:	2.02
// Date:		07.Jan.94
// Author:	Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date			By	Action
// 16.Aug.94	GG	Revised to use Visual Basic 
//			input dialog.
// 21.Nov.94	GG	Revised to use DRAFIX 3.0
// 18.Oct.95	GG	Revised to use arguments 
// 30.Sep.96	GG	Removed Heel styles 

    Display ("cursor", "wait", "Loading");

HANDLE
	hDDE,
	hChan,
	hTitle,			// TITLEBOX
	hWHBody,		// WHBODYBOX
	hLeg;			// LEGBOX 

STRING
	sPathJOBST,
	sFileNo,
	sLeg,
	sTmp,
	sTapeLengths;

// Clear User any user selections
//
   UserSelection ("clear") ;

// Find JOBST installed directory
//
   sPathJOBST =  GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;

   sLeg = GetArgument (1) ;

Print ("\n", sLeg) ;

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


// Get legbox handle
//
// Add this field just in case 
   Table("add", "field", "Leg","string");

   sTmp = "DB SymbolName = 'waistleg' AND DB fileno = '" + sFileNo + "' AND  DB Leg = '" + sLeg + "'" ;
   hChan = Open ("selection", sTmp);
   hLeg = %zero ;
   if (hChan) {
	ResetSelection(hChan);
   	hLeg = GetNextSelection(hChan) ;
	Close("selection", hChan); 
	}

// Start the application and get DDE link
//
   Execute ("application", sPathJOBST  + "\\waist\\whlegdia", "normal");
   hDDE = Open ("dde", "whlegdia", "whlegdia");

   if (hLeg && hDDE) {
  	GetDBData( hLeg, "TapeLengthsPt1", &sTapeLengths); 
  	GetDBData( hLeg, "TapeLengthsPt2", &sTmp); 
		sTapeLengths = sTapeLengths + sTmp ;
	Poke ( hDDE, "txtTapeLengths", sTapeLengths );

  	GetDBData( hLeg, "AnkleTape", &sTmp);
	Poke ( hDDE, "txtAnkleTape", sTmp);

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

   	GetDBData( hLeg, "Data", &sTmp);
	Poke ( hDDE, "txtData", sTmp);

	Poke ( hDDE, "txtUidLeg", MakeString("long",UID("get",hLeg)));
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

	Poke ( hDDE, "txtUidTitle", MakeString("long",UID("get",hTitle)));
  	}

    Close ( "dde", hDDE);

 Exit (%ok, " Leg details transfered");




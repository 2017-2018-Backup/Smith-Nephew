// File Name:	WH_FIGUR.D
// Purpose:	Figure the ankle MMs for both legs
//		Allow selection of fabric and templates
//
// Version:	2.01
// Date:		31.Jan.94
// Author:	Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
// 16.Nov.94	GG	Revised to use Visual Basic 
//			input dialog.
// 17.Nov.94	GG	JOBSTEX fabric
// 20.Nov.94	GG	DRAFIX 3.0 Modifications

   Display ("cursor", "wait", "Loading");

HANDLE
	hDDE,
	hChan,
	hTitle,			// TITLEBOX
	hWHBody,		// WHBODYBOX
	hLeftLeg,
	hRightLeg;		// LEGBOX 

STRING
	sPathJOBST,
	sFileNo,
	sLeg,
	sTmp,
	s1, s2, s3, s4, s5, s6, s7, s8, s9, s10,
	s11, s12, s13, s14, s15, s16, s17, s18, s19, s20;

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
	Exit(%cancel, "Patient details cannot be found\nPlease ensure that TITLEBOX has been used\nThen try again\n");
	}
   else
   	GetDBData( hTitle, "fileno", &sFileNo);

   Close("selection", hChan);

// Get "waistbox" handle
//
   sTmp = "DB SymbolName = 'waistbody' AND DB fileno = '" + sFileNo + "'" ;
   hChan = Open ("selection", sTmp);
   hWHBody = %zero ;
   if (hChan) {
	ResetSelection(hChan);
   	hWHBody = GetNextSelection(hChan) ;
	Close("selection", hChan); 
	}

// Get left legbox handle
//
   sLeg = "Left";
   sTmp = "DB SymbolName = 'waistleg' AND DB fileno = '" + sFileNo + "' AND  DB Leg = '" + sLeg + "'" ;
   hChan = Open ("selection", sTmp);
   hLeftLeg = %zero ;
   if (hChan) {
	ResetSelection(hChan);
   	hLeftLeg = GetNextSelection(hChan) ;
	Close("selection", hChan); 
	}


// Get RIGHT legbox handle
//
   sLeg = "Right";
   sTmp = "DB SymbolName = 'waistleg' AND DB fileno = '" + sFileNo + "' AND  DB Leg = '" + sLeg + "'" ;
   hChan = Open ("selection", sTmp);
   hRightLeg = %zero ;
   if (hChan) {
	ResetSelection(hChan);
   	hRightLeg = GetNextSelection(hChan) ;
	Close("selection", hChan); 
	}

// Start the application and get DDE link
//
   Execute ("application", sPathJOBST + "\\waist\\whfigure", "normal");
   hDDE = Open ("dde", "whfigure", "whfigure");

   if (hLeftLeg && hDDE) {
  	GetDBData( hLeftLeg, "TapeLengthsPt1", &s1); 
  	GetDBData( hLeftLeg, "TapeLengthsPt2", &s2); 
		sTmp = s1 + s2 ;
	Poke ( hDDE, "txtLeftLengths", sTmp );

  	GetDBData( hLeftLeg, "TapeMMs", &s3); 
  	GetDBData( hLeftLeg, "TapeMMs2", &s4); 
		sTmp = s3 + s4 ;
	Poke ( hDDE, "txtLeftMMs", sTmp );

  	GetDBData( hLeftLeg, "Grams", &s5); 
  	GetDBData( hLeftLeg, "Grams2", &s6); 
		sTmp = s5 + s6 ;
	Poke ( hDDE, "txtLeftStr", sTmp );

  	GetDBData( hLeftLeg, "Reduction", &s7); 
  	GetDBData( hLeftLeg, "Reduction2", &s8); 
		sTmp = s7 + s8 ;
	Poke ( hDDE, "txtLeftRed", sTmp );

  	GetDBData( hLeftLeg, "AnkleTape", &s9);
	Poke ( hDDE, "txtLeftAnkle", s9);

  	GetDBData( hLeftLeg, "Pressure", &s10);
	Poke ( hDDE, "txtLeftTemplate", s10);

	Poke ( hDDE, "txtUidLeftLeg", MakeString("long",UID("get",hLeftLeg)));
 	}

   if (hRightLeg && hDDE) {
  	GetDBData( hRightLeg, "TapeLengthsPt1", &s11); 
  	GetDBData( hRightLeg, "TapeLengthsPt2", &s12); 
		sTmp = s11 + s12 ;
	Poke ( hDDE, "txtRightLengths", sTmp );

  	GetDBData( hRightLeg, "TapeMMs", &s13); 
  	GetDBData( hRightLeg, "TapeMMs2", &s14); 
		sTmp = s13 + s14 ;
	Poke ( hDDE, "txtRightMMs", sTmp );

  	GetDBData( hRightLeg, "Grams", &s15); 
  	GetDBData( hRightLeg, "Grams2", &s16); 
		sTmp = s15 + s16 ;
	Poke ( hDDE, "txtRightStr", sTmp );

  	GetDBData( hRightLeg, "Reduction", &s17); 
  	GetDBData( hRightLeg, "Reduction2", &s18); 
		sTmp = s17 + s18 ;
	Poke ( hDDE, "txtRightRed", sTmp );

  	GetDBData( hRightLeg, "AnkleTape", &s19);
	Poke ( hDDE, "txtRightAnkle", s19);

  	GetDBData( hRightLeg, "Pressure", &s20);
	Poke ( hDDE, "txtRightTemplate", s20);

	Poke ( hDDE, "txtUidRightLeg", MakeString("long",UID("get",hRightLeg)));
 	}

   if (hWHBody && hDDE) {
  	GetDBData( hWHBody, "Fabric", &sTmp); 
	Poke ( hDDE, "txtFabric", sTmp);

  	GetDBData( hWHBody, "ID", &sTmp); 
	Poke ( hDDE, "txtLegStyle", sTmp);

	Poke ( hDDE, "txtUidBody", MakeString("long",UID("get",hWHBody)));
 	}


    if (hTitle && hDDE) {

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

 Exit (%ok, "FIGURE details transfered");

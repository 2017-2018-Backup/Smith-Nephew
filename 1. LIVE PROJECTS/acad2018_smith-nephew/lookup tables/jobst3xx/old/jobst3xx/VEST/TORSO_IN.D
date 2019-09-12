// File Name:	TORSO_IN.D
//
// Purpose:	Starts the VB programme that is used to 
//		draw the TORSO Band
//
// Version:	1.01
//
// Date:	13.Jan.98
//
// Author:	Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
// 04.Jun.98		GG	Revised heights
//				Chest to waist and chest to EOS
//
//---------------------------------------------------------------------------------------------
// Notes:
// 

Display ("cursor", "wait", "Loading Data");

HANDLE	
	hChan,
	hSym,
	hDDE,
	hEnt;
	
XY	
	xySymO,
	xyScale;
	
ANGLE
	aAngle;

STRING	
	sSymName,
	sSymSelection,
	sPathJOBST,
	sWorkOrder,
	sSide,
	// Strings are initialized to empty "", this ensures empty strings
	// are poked if there is no DB field.
	s1, s2, s3, s4, s5, s6, s7, 
	s10, s11, s12, s13, s14, s15, s16, s17, s18, s19,
	s29;

// Find JOBST installed directory
//
   sPathJOBST =  GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;

// Start the application and get DDE link
//
   Execute ("application", sPathJOBST + "\\VEST\\torsodia", "normal");
   hDDE = Open ("dde", "torsodia", "torsodia");
   if ( !hDDE) Exit (%abort, "Unable to open \"torsodia\" dialogue!") ;

// Create new fields
// Note If the field aready exists this is ignored
   Table("add" , "field", "SLgButt", "string");
   Table("add" , "field", "SFButt", "string");

   

// Get all of the symbols for uploading
//
   sSymSelection = "DB SymbolName = 'mainpatientdetails' OR " + "DB SymbolName = 'vestbody'"  ;

   hChan = Open ("selection", sSymSelection );  
   if (hChan && hDDE ) {
      	ResetSelection(hChan);
	// Get Symbol data base fields
	//
          	while (hSym = GetNextSelection(hChan)) {
   		GetGeometry(hSym, &sSymName, &xySymO, &xyScale, &aAngle);
		 if(StringCompare(sSymName, "mainpatientdetails" )){ 
			GetDBData( hSym, "fileno", &s1); Poke ( hDDE, "txtFileNo", s1);
			GetDBData( hSym, "patient", &s2); Poke ( hDDE, "txtPatientName", s2);
 			GetDBData( hSym, "age",  &s3); Poke ( hDDE, "txtAge", s3);
			GetDBData( hSym, "units", &s4); Poke ( hDDE, "txtUnits",s4);
			GetDBData( hSym, "sex", &s5); Poke ( hDDE, "txtSex",s5);
			GetDBData( hSym, "Diagnosis", &s6); Poke ( hDDE, "txtDiagnosis", s6);
			GetDBData( hSym, "WorkOrder", &s7); Poke ( hDDE, "txtWorkOrder" ,s7);
			Poke ( hDDE, "txtUidMPD", MakeString("long",UID("get",hSym)));
   			}
		if(StringCompare(sSymName, "vestbody" )){ 
//			GetDBData( hSym, "LtSCir", &s10);	Poke ( hDDE, "txtCir(0)", s10);
//			GetDBData( hSym, "RtSCir", &s11);	Poke ( hDDE, "txtCir(1)", s11);
//			GetDBData( hSym, "NeckCir", &s12);	Poke ( hDDE, "txtCir(2)", s12);
//			GetDBData( hSym, "SWidth", &s13);	Poke ( hDDE, "txtCir(3)", s13);	  
//			GetDBData( hSym, "S_Waist", &s14);	Poke ( hDDE, "txtCir(4)", s14);
			// NB The DB variable used below is re-used (the name is not significant)
			GetDBData( hSym, "SLgButt", &s14);		Poke ( hDDE, "txtCir(4)", s14);		// Chest to waist
			GetDBData( hSym, "ChestCir", &s15);	Poke ( hDDE, "txtCir(5)", s15);
			GetDBData( hSym, "WaistCir", &s16 );	Poke ( hDDE, "txtCir(6)", s16);
//			GetDBData( hSym, "S_EOS", &s17);	Poke ( hDDE, "txtCir(7)", s17);
			// NB The DB variable used below is re-used (the name is not significant)
			GetDBData( hSym, "SFButt", &s17);	Poke ( hDDE, "txtCir(7)", s17);		// Chest to EOS 
			GetDBData( hSym, "EOSCir", &s18);	Poke ( hDDE, "txtCir(8)", s18);
			GetDBData( hSym, "Fabric", &s19); 	Poke ( hDDE, "txtCombo(10)", s19);
			GetDBData( hSym, "Closure", &s29);		Poke ( hDDE, "txtCombo(9)", s29);
//			GetDBData( hSym, "WaistCirUserFac", &s30);	Poke ( hDDE, "txtCombo(0)", s30);
//			GetDBData( hSym, "EOSCirUserFac", &s31);	Poke ( hDDE, "txtCombo(1)", s31);
			Poke ( hDDE, "txtUidVB", MakeString("long",UID("get",hSym)));
   			}
   		}
	}
   
   Close ( "selection", hChan); 
   Close ( "dde", hDDE);



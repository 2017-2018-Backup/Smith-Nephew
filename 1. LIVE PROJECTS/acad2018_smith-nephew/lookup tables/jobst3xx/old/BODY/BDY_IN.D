// File Name:	BDY_IN.D
//
// Purpose:	Starts the VB programme that is used to 
//		draw the BodySuit
//
// Version:	1.00
//
// Date:		7/5/96
//
// Author:		Paul O'Rawe
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
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
	s20, s21, s22, s23, s24, s25, s26, s27,	s28, s29, 
	s30, s31, s32, s33, s34, s35, s36,
	s40, s41, s42, s43, s44, s45, s46, s47, s48, s49;

// Find JOBST installed directory
//
   sPathJOBST =  GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;

// Start the application and get DDE link
//
   Execute ("application", sPathJOBST + "\\BODY\\BODYSUIT", "normal");
   hDDE = Open ("dde", "bodysuit", "bodysuit");
   if ( !hDDE) Exit (%abort, "Unable to open \"bodysuit\" dialogue!") ;

// Get all of the symbols for uploading
//
   sSymSelection = "DB SymbolName = 'mainpatientdetails' OR " + "DB SymbolName = 'suitbody'"  ;

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
		if(StringCompare(sSymName, "suitbody" )){ 
			GetDBData( hSym, "LtSCir", &s10);	Poke ( hDDE, "txtCir(0)", s10);
			GetDBData( hSym, "RtSCir", &s11);	Poke ( hDDE, "txtCir(1)", s11);
			GetDBData( hSym, "NeckCir", &s12);	Poke ( hDDE, "txtCir(2)", s12);
			GetDBData( hSym, "SWidth", &s13);	Poke ( hDDE, "txtCir(3)", s13);	  
			GetDBData( hSym, "S_Waist", &s14);	Poke ( hDDE, "txtCir(4)", s14);
			GetDBData( hSym, "ChestCir", &s15);	Poke ( hDDE, "txtCir(5)", s15);
			GetDBData( hSym, "WaistCir", &s16 );	Poke ( hDDE, "txtCir(6)", s16);
			GetDBData( hSym, "S_Breast", &s19);	Poke ( hDDE, "txtCir(9)", s19);
			GetDBData( hSym, "BreastCir", &s20);	Poke ( hDDE, "txtCir(10)", s20);
			GetDBData( hSym, "NippleCir", &s21);	Poke ( hDDE, "txtCir(11)", s21);
			GetDBData( hSym, "Fabric", &s22); 	Poke ( hDDE, "txtCombo(10)", s22);
			GetDBData( hSym, "BraLtCup", &s23);	Poke ( hDDE, "txtCombo(3)", s23);
			GetDBData( hSym, "BraRtCup", &s24);	Poke ( hDDE, "txtCombo(4)", s24);
			GetDBData( hSym, "BraRtDisk", &s25);	Poke ( hDDE, "txtRightDisk", s25);
			GetDBData( hSym, "BraLtDisk", &s26);	Poke ( hDDE, "txtLeftDisk", s26);
			GetDBData( hSym, "LtAxillaType", &s27);	Poke ( hDDE, "txtCombo(5)", s27);
			GetDBData( hSym, "RtAxillaType", &s28);	Poke ( hDDE, "txtCombo(6)", s28);
			GetDBData( hSym, "Closure", &s29);		Poke ( hDDE, "txtCombo(9)", s29);
			GetDBData( hSym, "ChestCirUserFac", &s31);	Poke ( hDDE, "txtCombo(1)", s31);
			GetDBData( hSym, "WaistCirUserFac", &s30);	Poke ( hDDE, "txtCombo(0)", s30);
			GetDBData( hSym, "BreastCirUserFac", &s32);	Poke ( hDDE, "txtCombo(2)", s32);
			GetDBData( hSym, "NeckType", &s33);		Poke ( hDDE, "txtCombo(7)", s33);
			GetDBData( hSym, "BackNeckType", &s34);	Poke ( hDDE, "txtCombo(8)", s34);
			GetDBData( hSym, "NeckDimension", &s35);  	Poke ( hDDE, "txtFrontNeck", s35);
			GetDBData( hSym, "BackNeckDim", &s36);	Poke ( hDDE, "txtBackNeck", s36);

			// Extra Fields
			GetDBData( hSym, "SFButt", &s40);	Poke ( hDDE, "txtCir(12)", s40);
			GetDBData( hSym, "SLgButt", &s41);	Poke ( hDDE, "txtCir(13)", s41);
			GetDBData( hSym, "LgButtCir", &s42);	Poke ( hDDE, "txtCir(14)", s42);
			GetDBData( hSym, "LgButtCirUserFac", &s47);	Poke ( hDDE, "txtCombo(11)", s47);
			GetDBData( hSym, "LtThCir", &s43);	Poke ( hDDE, "txtCir(15)", s43);
			GetDBData( hSym, "ThCirUserFac", &s48);	Poke ( hDDE, "txtCombo(12)", s48);
			GetDBData( hSym, "RtThCir", &s44);	Poke ( hDDE, "txtCir(16)", s44);
			GetDBData( hSym, "LegStyle", &s45);	Poke ( hDDE, "txtLegStyle", s45);
			GetDBData( hSym, "CrotchStyle", &s46); Poke ( hDDE, "txtCrotchStyle", s46);

			Poke ( hDDE, "txtUidVB", MakeString("long",UID("get",hSym)));
   			}
   		}
	}
   
   Close ( "selection", hChan); 
   Close ( "dde", hDDE);



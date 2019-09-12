// File Name:	WEB_INR.D
//
// Purpose:	Starts the VB programme that is used to 
//		draw the Web Spacer for the 
//		RIGHT Glove
//
// Version:	1.01
//
// Date:		22.August.96
//
// Author:		Gary George
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
	s1, s2, s3, s4, s5, s6, s10, s11,s12, s20, s21, s22, s23, s24,s25, s26, s27,s28,s29;	

// Set side
//
   sSide = "Right";
   
// Find JOBST installed directory
//
   sPathJOBST =  GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;

// Start the application and get DDE link
//
   Execute ("application", sPathJOBST + "\\WEBSPACR\\WEBSPACR", "normal");
   hDDE = Open ("dde", "webspacr", "webspacr");
   if ( !hDDE) Exit (%abort, "Unable to open \"webspacr\" drawing routine!") ;

// Get all of the symbols for uploading
//
   sSymSelection = 	"DB SymbolName = 'mainpatientdetails' OR " +
   		"(DB SymbolName = 'gloveglove' AND DB Sleeve ='"  + sSide +"') OR " +
   		"DB SymbolName = 'glovecommon'"  ;

   hChan = Open ("selection", sSymSelection );  
   if (hChan && hDDE ) {
      	ResetSelection(hChan);
      	// Poke side
      	//
      	Poke ( hDDE, "txtSide", sSide);
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
			GetDBData( hSym, "WorkOrder", &s10); Poke ( hDDE, "txtWorkOrder",s10);
			Poke ( hDDE, "txtUidMPD", MakeString("long",UID("get",hSym)));
   			}
      		if(StringCompare(sSymName, "glovecommon" )){ 
			GetDBData( hSym, "Fabric", &s11); Poke ( hDDE, "txtFabric",s11);
			Poke ( hDDE, "txtUidGC", MakeString("long",UID("get",hSym)));
   			}
      		if(StringCompare(sSymName, "gloveglove" )){ 
			GetDBData( hSym, "TapeLengths", &s20); Poke ( hDDE, "txtTapeLengths",s20);
			GetDBData( hSym, "TapeLengths2", &s21);  Poke ( hDDE, "txtTapeLengths2",s21);	
			GetDBData( hSym, "TapeLengthPt1", &s22); Poke ( hDDE, "txtTapeLengthPt1",s22);			
			GetDBData( hSym, "Data", &s26); Poke ( hDDE, "txtDataGlove",s26);			
			GetDBData( hSym, "WristPleat", &s27); Poke ( hDDE, "txtWristPleat",s27);			
			GetDBData( hSym, "ShoulderPleat", &s28); Poke ( hDDE, "txtShoulderPleat",s28);			
			Poke ( hDDE, "txtUidGlove", MakeString("long",UID("get",hSym)));
   			}
   		}
	}
   
   Close ( "selection", hChan); 
   Close ( "dde", hDDE);
 
   Exit (%ok, "Patient and Glove Details transfered");


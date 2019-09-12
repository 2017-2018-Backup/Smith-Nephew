// File Name:	CL_PDWOD.D
//
// Purpose:	Starts the VB programme that is used to 
//		Display and Modify the Current
//		Work Order and Patient Details
//
// Version:	1.01
//
// Date:		26.Jun.95
//
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
//---------------------------------------------------------------------------------------------
// Notes:
// 	Transfers existing patient data from "mainpatientdetails"
//	and existing workorder details from common symbols.
//
//	This MACRO is essentially the same as CL_TXF.D
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
	sCurrentCADFile, 
	sWorkOrder, sWO_TXF,
	// Strings are initialized to empty "", this ensures empty strings
	// are poked if there is no DB field.
	s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s20, s30, s40, s50, s60, s70;	

// Find JOBST installed directory
//
   sPathJOBST =  GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;

// Check if JOBSTNEW needs to be run
//
   SetData("PathMacro", sPathJOBST);

   if ( Table("find", "layer", "titlebox") == %badtable) { @JOBSTNEW.D; }
   
// Start the application and get DDE link
// Note use of CLTXFDIA VB programme.
//
   Execute ("application", sPathJOBST + "\\CADLINK\\CLTXFDIA", "normal");
   hDDE = Open ("dde", "cltxfdia", "cltxfdia");
   if ( !hDDE) Exit (%abort, "Unable to open \"cltxfdia\" dialogue!") ;

// Get "mainpatientdetails" uploading
//

   hChan = Open ("selection",	"DB SymbolName = 'mainpatientdetails'"  );  
   if (hChan && hDDE ) {
      	ResetSelection(hChan);
      	Poke ( hDDE, "txtInvokedFrom", "DRAFIX");
      	
      	GetData ("DocFile", &sCurrentCADFile);
 	Poke ( hDDE, "txtCurrentCADFile", sCurrentCADFile ); // Current CAD file
 	Poke ( hDDE, "txtInitialCADFile", sCurrentCADFile );    // In this case Inital = Current	

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
			GetDBData( hSym, "WorkOrder", &s7); Poke ( hDDE, "txtMPDwo",s7);	
	   		GetDBData( hSym, "orderdate", &s8); Poke ( hDDE, "txtOrderDate",s8);
   			GetDBData( hSym, "TemplateEngineer", &s9); Poke ( hDDE, "txtTemplateEngineer",s9);
			Poke ( hDDE, "txtUidMPD", MakeString("long",UID("get",hSym)));
   			}
   		}
	}
   
   Close ( "selection", hChan); 
   Close ( "dde", hDDE);
 
   Exit (%ok, "Patient Details transfered");


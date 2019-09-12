// File Name:	CL_TXF.D
//
// Purpose:	Starts the VB programme that is used to 
//		initialise a drawing with the data transfered
//		from TRITON 
//
// Version:	1.01
//
// Date:		09.Jun.95
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
//	The VB programme can then load the TXF (Triton eXchange File)
//	and intellegently deal with problems

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
	sWO_TXF, 
	sCurrentCADFile, 
	sInitialCADFile,
	// Strings are initialized to empty "", this ensures empty strings
	// are poked if there is no DB field.
	s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s20, s30, s40, s50, s60, s70;	

// Start the application and get DDE link
// Find JOBST installed directory
//
   sPathJOBST =  GetProfileString("JOBST", "PathJOBST", "\\JOBST", "DRAFIX.INI") ;

   Execute ("application", sPathJOBST + "\\CADLINK\\CLTXFDIA", "normal");
   hDDE = Open ("dde", "cltxfdia", "cltxfdia");
   if ( !hDDE) Exit (%abort, "Unable to open \"cltxfdia\" dialogue!") ;

// Get all of the symbols for uploading
//
   sSymSelection = 	"DB SymbolName = 'mainpatientdetails'" ;

   hChan = Open ("selection", sSymSelection );  
   if (hChan && hDDE ) {
      	ResetSelection(hChan);
      	// Get Triton eXchange File and Work Order
      	//
 	Poke ( hDDE, "txtInvokedFrom", "imageABLE");
 	GetData ( "UserChar1", &sWO_TXF);		// TXF File
 	Poke ( hDDE, "txtWO_TXF", sWO_TXF);
 	
 	GetData ( "UserChar2", &sWorkOrder);		// Work order	
 	Poke ( hDDE, "txtWorkOrder", sWorkOrder);
 	
 	GetData ( "UserChar3", &sInitialCADFile );	// The CAD file the Drawing will 
 	Poke ( hDDE, "txtInitialCADFile", sInitialCADFile );// be based on or "NEW" or "INSPECTION"
 	
   	GetData ("DocFile", &sCurrentCADFile);		// The current cad file
 	Poke ( hDDE, "txtCurrentCADFile", sCurrentCADFile );
 	
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
 
   Exit (%ok, "Patient and Work Order Details transfered");


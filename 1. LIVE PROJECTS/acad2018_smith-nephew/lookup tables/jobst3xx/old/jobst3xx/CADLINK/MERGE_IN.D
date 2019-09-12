// File Name:	MERGE_IN.D
//
// Purpose:		Starts the VB programme that is used to 
//			Select the CAD file to merge with the current drawing
// Version:		1.01
//
// Date:		06.Feb.98
//
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
//---------------------------------------------------------------------------------------------
// Notes:


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

// Start the application and get DDE link
// Note use of CLTXFDIA VB programme.
//
   Execute ("application", sPathJOBST + "\\CADLINK\\merge", "normal");
   hDDE = Open ("dde", "merge", "merge");
   if ( !hDDE) Exit (%abort, "Unable to open \"merge\" dialogue!") ;

// Get "mainpatientdetails" uploading
//

   hChan = Open ("selection",	"DB SymbolName = 'mainpatientdetails'"  );  
   if (hChan && hDDE ) {
      	ResetSelection(hChan);
      	
      	GetData ("DocFile", &sCurrentCADFile);
 	Poke ( hDDE, "txtCurrentCADFile", sCurrentCADFile ); // Current CAD file

 	// Get Symbol data base fields
	//
          	while (hSym = GetNextSelection(hChan)) {
   		GetGeometry(hSym, &sSymName, &xySymO, &xyScale, &aAngle);
      		if(StringCompare(sSymName, "mainpatientdetails" )){ 
			GetDBData( hSym, "fileno", &s1); Poke ( hDDE, "txtFileNo", s1);
			GetDBData( hSym, "patient", &s2); Poke ( hDDE, "txtPatientName", s2);
			GetDBData( hSym, "WorkOrder", &s7); Poke ( hDDE, "txtWorkOrder",s7);	
	   		GetDBData( hSym, "orderdate", &s8); Poke ( hDDE, "txtOrderDate",s8);
   			}
   		}
	}
   
   Close ( "selection", hChan); 
   Close ( "dde", hDDE);
 
   Exit (%ok, "Patient Details transfered");


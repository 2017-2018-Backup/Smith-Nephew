// File Name:	STMPSTAM.D
// Purpose:	To start the stamp VB routine
//		General Start.
//
// Version:	1.1
// Date:		4.Aug.94
// Author:		Gary George
//

HANDLE
	hDDE;
STRING	
	sStampType,
	sPathJOBST,
	sPath;

// Find JOBST installed directory
//
   sPathJOBST =  GetProfileString("JOBST", "PathJOBST", "C:\\JOBST", "DRAFIX.INI") ;

// Initalise
//
   sStampType = "Arm";

// Start the application and get link, poke Path and required Stamp type
//
    Execute ("application", sPathJOBST + "\\stamps\\stamp", "normal");
    hDDE = Open ("dde", "stamp", "stamp");

    if ( hDDE) {
	Poke ( hDDE, "txtStampType", sStampType);
	Close ( "dde", hDDE);
	}

// Finished 
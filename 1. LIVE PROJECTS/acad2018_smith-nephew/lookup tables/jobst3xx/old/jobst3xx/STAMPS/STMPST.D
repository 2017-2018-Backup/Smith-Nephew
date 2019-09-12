// File Name:	STMPST.D;
// Purpose:	To start the stamp VB routine
//
// Version:	1.1
// Date:		4.Aug.94
// Author:		Gary George
//

HANDLE
	hDDE;
STRING	
	sStampType,
	sPath;

// Initalise
//
   sStampType = "Waist Ht";

// Check if a link to an existing instance of the "stamp" VB program can be made.
// If it can then Close it. As we don't know the state of this instance.

    hDDE = Open ("dde", "stamp", "Stamp");
    if ( hDDE) Close ("dde", hDDE) ;

// Start the application and get link, poke Path and required Stamp type
//
    Execute ("application", "c:\\jobst\\stamps\\stamp", "normal");
    hDDE = Open ("dde", "stamp", "Stamp");

    if ( hDDE) {
	Poke ( hDDE, "txtStampType", sStampType);
	Close ( "dde", hDDE);
	}

// Finished 
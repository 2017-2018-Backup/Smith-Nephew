// File Name:	ZIPSELCT.D
// Purpose:	Used to select all the enities that are part 
//			of a Zipper
//
// Version:	1.01
// Date:		20.Sep.95
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date	By	Action
//---------------------------------------------------------------------------------------------
//
// 

HANDLE
	hEnt,
	hChan;
	
STRING 
	sEntityID,
	sTmp;

NUMBER
	nFound;
	
// Clear current user selection
//
   UserSelection("Clear");
   UserSelection("Update");

// Prompt user to select part of a Zipper
//
   if (! hEnt=GetEntity ("Select Zipper", "all")) Exit (%ok, "User Cancelled") ;

   Display ("cursor", "wait", "Searching for more data");
   GetDBData (hEnt, "ID", &sEntityID);
 
   if (StringLength (sEntityID) <=0) Exit (%ok, "Selected entity does not have an ID field");
   sTmp = "DB ID = '" + sEntityID + "'" ;
   nFound = 0 ;
   hChan = Open ("selection", sTmp );
   if (hChan) {
	ResetSelection(hChan);
	while (hEnt = GetNextSelection(hChan)) {
		// Only select if "Zipper" DB value is true 
		sTmp = "" ;
		GetDBData(hEnt,"Zipper", &sTmp) ;
		if (StringCompare( sTmp, "1")){
			nFound = nFound + 1;
			UserSelection("add", hEnt) ;
			}
		}
   	}
   Close("selection", hChan); 
     
// Exit if nothing found
//
   if (nFound == 0 ) Exit (%ok, "Selected entity does not belong to a Zipper");
   
   UserSelection ("Update") ;
   
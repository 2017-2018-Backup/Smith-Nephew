// File Name:	ARMVRBOD.D
// Purpose:	To read data from the BODYBOX symbol
//		And poke it into the Visual Basic program
//
// Version:	1.01
// Date:		02.Nov.94
// Author:		Gary George
// Called By:	ARML.D and ARMR.D
//   
// -----------------------------------------------------------------------------------------
// REVISIONS
// date		by	action
// -----------------------------------------------------------------------------------------


// Variables declared in ARML.D or ARMR.D 
// STRING	   sData, sTmp, sFileNo, sSleeve;
// HANDLE  hBody, hDDE, hChan;
// sSleeve, sFileNo and hDDE are set in ARML.D or ARMR.D 

   sData = "None 0 0 0";	//Initialise to None

// Get BODYBOX Details
//
   hChan = Open ("selection", "DB SymbolName = 'vestbody' AND DB fileno = '" + sFileNo + "'");
   ResetSelection(hChan);
   hBody = GetNextSelection(hChan);
   if (hBody) {
   	if (StringCompare(sSleeve, "Right")) {
		// If Left and Right axilla are of different heights then there is a specific
		// set of values for the Right.  If these are found they are used.
		//
		GetDBData ( hBody,  "AFNRadRight", &sTmp) ;
		if ( !Value ( "scalar", sTmp )) GetDBData ( hBody,  "AxillaFrontNeckRad", &sTmp) ;
		sData = sTmp ;
		
		GetDBData ( hBody,  "ABNRadRight", &sTmp) ;
		if ( !Value ( "scalar", sTmp )) GetDBData ( hBody,  "AxillaBackNeckRad", &sTmp) ;
		sData = sData + " " + sTmp ;		
		
		GetDBData ( hBody,  "SBRaglanRight", &sTmp) ;
		if ( !Value ( "scalar", sTmp )) GetDBData ( hBody,  "ShoulderToBackRaglan", &sTmp) ;
		sData = sData + " " + sTmp ;		
		
		GetDBData( hBody, "RtAxillaType", &sTmp);
		sData = sTmp + " " + sData  ;		
		}
	else {
		GetDBData ( hBody,  "AxillaFrontNeckRad", &sTmp) ;
		sData = sTmp ;

    		GetDBData ( hBody,  "AxillaBackNeckRad", &sTmp) ;
		sData = sData + " " + sTmp ;		

    		GetDBData ( hBody,  "ShoulderToBackRaglan", &sTmp) ;
		sData = sData + " " + sTmp ;		

		GetDBData( hBody, "LtAxillaType", &sTmp);
		sData = sTmp + " " + sData  ;		
		}
	}

   Poke ( hDDE, "txtVestRaglan", sData) ;
  
   Close ("selection", hChan);
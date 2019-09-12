// File Name:	WHLBLBOD.D
// Purpose:	Get Waist Height body details from Waist box
// Version:	1.01
// Date:		14.Mar.94
// Author:	Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
// 8.Nov.94	GG	nMaxCir added to "ID" Field
//
// Get details from waistbox 
//
   sTmp = "DB SymbolName = 'waistbody' AND DB fileno = '" + sFileNo + "'" ;
   hChan = Open ("selection", sTmp);
   hWHBody = %zero ;
   if (hChan) {
	ResetSelection(hChan);
   	hWHBody = GetNextSelection(hChan) ;
	Close("selection", hChan); 
	}

   if (hWHBody) {
   	GetDBValue( hWHBody, "LeftThighCir", &sLeftThighCir, &nLeftThighCir); 
 	   nLeftThighCir = FNRound ( FNDecimalise(nLeftThighCir) * nUnitsFac) ;
   	GetDBValue( hWHBody, "RightThighCir", &sRightThighCir, &nRightThighCir); 
	   nRightThighCir = FNRound ( FNDecimalise(nRightThighCir) * nUnitsFac) ;
   	GetDBData( hWHBody, "CrotchStyle", &sCrotchStyle); 
	GetDBData( hWHBody, "ID", &sLine); 
	ScanLine(sLine, "blank", &nLegStyle, &nLeftLegStyle, &nRightLegStyle, &nMaxCir);		
	}
   Close("selection", hChan); 

// Using details from waist box decide which leg is 1st leg
//

   if ( (nLeftThighCir - nRightThighCir) > 1) {
	sLeg = "Right" ;
	RightLeg = %true ;
	LeftLeg = %false ;
	}
   else	{
	sLeg = "Left" ;
	LeftLeg = %true ;
   	RightLeg = %false ;
	}

// Crotch type
//
   OpenCrotch = %false;
   ClosedCrotch = %false ;
   if (StringCompare("Open Crotch", sCrotchStyle))
	OpenCrotch = %true ;
   else 
	ClosedCrotch = %true ;

// Leg  type
//
   if (nLegStyle == 1)
	PantyLeg = %true ;
   else 
	PantyLeg = %false ;


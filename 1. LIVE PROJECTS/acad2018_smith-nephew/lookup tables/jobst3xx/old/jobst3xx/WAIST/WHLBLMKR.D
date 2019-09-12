// File Name:	WHLBLMKR.D
// Purpose:	Get Body Points from 1st leg markers
//
// Version:	1.01
// Date:		14.Mar.94
// Author:	Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action


// Get marker details
//
   LeftCO_CenterArrowFound = %false;
   sCO_LargestTopID   =  sFileNo + sLeg+ "CO_LargestTop";
   sCO_LargestBottID   = sFileNo + sLeg+ "CO_LargestBott";

   if (nLegStyle == 4 )
   	sCO_CenterArrowID  = sFileNo + "Right" + "CO_CenterArrow"; // Label Right for W4OC Right
   else
   	sCO_CenterArrowID  = sFileNo + "Left" + "CO_CenterArrow"; // Label Left Crotch only

   sTmp = "type = 'Marker' AND (DB ID ='" + sCO_CenterArrowID + "' OR DB ID ='" + sCO_LargestTopID + "' OR DB ID ='" + sCO_LargestBottID + "')";

   hChan = Open ("selection", sTmp);
   if (hChan) {
	ResetSelection(hChan);
	while (hEnt = GetNextSelection(hChan)) {
		GetDBData(hEnt,"ID", &sTmp) ;
		GetGeometry(hEnt, &sType, &xyPt1, &xyTmp, &aAngle) ;
	          if ( xyPt1.x != 0 || xyPt1.y != 0) {
		if (StringCompare(sTmp, sCO_CenterArrowID)) {
			GetGeometry(hEnt, &sTmp, &xyCO_CenterArrow, &xyTmp, &aAngle) ;
			GetDBData(hEnt,"Data", &sTmp) ;
   			LeftCO_CenterArrowFound = %true;
			nOpenOff = Value ("scalar", sTmp);
			}
		if (StringCompare(sTmp,  sCO_LargestTopID))
			GetGeometry(hEnt, &sTmp, &xyCO_LargestTop, &xyTmp, &aAngle) ;
		if (StringCompare(sTmp,  sCO_LargestBottID))
			GetGeometry(hEnt, &sTmp, &xyCO_LargestBott, &xyTmp, &aAngle) ;
			}
		}
	}
   Close("selection", hChan); 

   if (LeftCO_CenterArrowFound == %false){
	 if (nLegStyle == 4)
		Exit (%cancel, "RIGHT leg crotch not found to label!\n\n") ;
	 else
		Exit (%cancel, "LEFT leg crotch not found to label!\n\n") ;
	}






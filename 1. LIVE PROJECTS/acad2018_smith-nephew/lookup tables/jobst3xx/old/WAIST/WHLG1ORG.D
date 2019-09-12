// File Name:	WHLG1ORG.D
// Purpose:	Establish the origin point to draw the leg
//		and leg template
//	
// Version:	1.02
// Date:		11.Jul.94
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
//---------------------------------------------------------------------------------------------
// 22.Dec.94	gg	Enabled brief (nLegStyle==2)
// Get marker details
//

   sThing =sFileNo+sFirstLeg ;
   BodyFound = %false ;
   xyFold.x = 10000 ;	// Dummy height
   hChan = Open ("selection", "type = 'Marker'");
   if (hChan) {
	ResetSelection(hChan);
	while (hEnt = GetNextSelection(hChan)) {
		GetDBData(hEnt,"ID", &sTmp) ;
		GetGeometry(hEnt, &sRef, &xyPt1, &xyTmp, &aAngle) ;
	          if ( xyPt1.x != 0 || xyPt1.y != 0) {
		if (StringCompare(sTmp,  sThing+"O")) {
			GetGeometry(hEnt, &sTmp, &xyOBody, &xyTmp, &aAngle) ;
			BodyFound = %true;
			}
		if (StringCompare(sTmp,  sThing+"TOS"))
			GetGeometry(hEnt, &sTmp, &xyTOS , &xyTmp, &aAngle) ;
		if (StringCompare(sTmp,  sThing+"Waist"))
			GetGeometry(hEnt, &sTmp, &xyWaist , &xyTmp, &aAngle) ;
		if (StringCompare(sTmp,  sThing+"MidPoint"))
			GetGeometry(hEnt, &sTmp, &xyMidPoint , &xyTmp, &aAngle) ;
		if (StringCompare(sTmp,  sThing+"Largest"))
			GetGeometry(hEnt, &sTmp, &xyLargest, &xyTmp, &aAngle) ;
		if (StringCompare(sTmp,  sThing+"Fold"))
			GetGeometry(hEnt, &sTmp, &xyFold, &xyTmp, &aAngle) ;
		if (StringCompare(sTmp,  sThing+"Thigh"))
			GetGeometry(hEnt, &sTmp, &xyThigh, &xyTmp, &aAngle) ;
		if (StringCompare(sTmp,  sThing+"FoldPanty"))
			GetGeometry(hEnt, &sTmp, &xyFoldPanty, &xyTmp, &aAngle) ;
			}
		}
	}
   Close("selection", hChan);
 
// Exit if no body origin found
   if (!BodyFound)
	Exit(%cancel, "No BODY found!, draw Cut-Out and try again.\n");   
   
// Establish leg start position relative to the body 
   nBodyLegTapePos = int( nFoldHt /1.5) + 6;
   if ( nFoldHt < nExtTempltTol1max && !FootLess ) 
	nLegStartHt = FNLegStartHt ( nFirstTape, nBodyLegTapePos) ;
   if (FootLess)
	nLegStartHt = FNLegStartHt ( nFirstTape, nBodyLegTapePos) ; // Panty
   if ( (nFoldHt >= nExtTempltTol2min && nFoldHt <= nExtTempltTol2max) && !FootLess ) 
	nLegStartHt = FNLegStartHt ( nFirstTape, nBodyLegTapePos - 1) ;
   if ( nFoldHt > nExtTempltTol3min && !FootLess) 
	nLegStartHt = FNLegStartHt ( nFirstTape, nBodyLegTapePos - 2) ;

   xyOtemplate.y = xyOBody.y;
   xyOtemplate.x =  xyOBody.x - nLegStartHt;
	
PRSetLayerConstruct();
hEnt= AddEntity("marker","xmarker",xyOtemplate,0.125);
if (hEnt) {
  sID=StringMiddle(MakeString("long",UID("get",hEnt)), 1, 4) ; 
  while (StringLength(sID) < 4) sID = sID + " ";
  sID = sID + sFileNo + sLeg ;
  SetDBData (hEnt,"ID",sID + "Origin");
  SetDBData (hEnt,"units", sUnits) ;
  if (FootLess)
  	 SetDBData (hEnt,"Data", "WHPT") ;
  else
  	 SetDBData (hEnt,"Data", "WHFT") ;  
  SetDBData (hEnt,"Pressure", sPressure) ;
  SetDBData (hEnt,"TapeLengthsPt1",StringMiddle(sTapeLengths, 1, 60));
  SetDBData (hEnt,"TapeLengthsPt2",StringMiddle(sTapeLengths, 61, StringLength(sTapeLengths)-60));
  SetDBData (hEnt,"patient", sPatient) ;
  SetDBData (hEnt,"age", sAge) ;
  SetDBData (hEnt,"WorkOrder", sWorkOrder) ;
//  SetDBData (hEnt,"TOSGivenRed", sTOSRed) ;
//  SetDBData (hEnt,"WaistGivenRed", sWaistRed) ;
//  SetDBData (hEnt,"FoldHt", sFoldHt) ;
//  SetDBData (hEnt,"WaistCir", sWaistCir) ;
//  SetDBData (hEnt,"WaistHt", sWaistHt) ;
//  SetDBData (hEnt,"TOSCir", sTOSCir) ;
//  SetDBData (hEnt,"TOSHt", sTOSHt) ;
  SetDBData (hEnt,"Body", sLegStyle) ;
  SetDBData (hEnt,"FirstTape", MakeString("scalar", nFirstTape));
  SetDBData (hEnt,"LastTape", MakeString("scalar", nLastTape));
  //Dummy fold so that we can find it when editing the leg
  hEnt= AddEntity("marker","xmarker",xyFold,0.125);
  SetDBData (hEnt,"ID",sID + "Fold");
  }
	
// File Name:	WHLG2ORG.D
// Purpose:	Establish the origin point to draw the leg
//		and leg template
//	
// Version:	1.01
// Date:		11.Jul.94
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
//---------------------------------------------------------------------------------------------
// Get marker details
//
   sThing =sFileNo+sFirstLeg ;
   BodyFound = %false ;
   xyFold.x = 10000 ;	// Dummy height

// setup values for tape
// Establish leg start position relative to the body 

   nBodyLegTapePos = int( nFoldHt /1.5) + 6;
   if ( nFoldHt < nExtTempltTol1max && !FootLess ) 
	nLegStartHt = FNLegStartHt ( nFirstTape, nBodyLegTapePos );
   if (FootLess)
	nLegStartHt = FNLegStartHt ( nFirstTape, nBodyLegTapePos ) ; // Panty
   if ( (nFoldHt >= nExtTempltTol2min && nFoldHt <= nExtTempltTol2max) && !FootLess ) 
	nLegStartHt = FNLegStartHt ( nFirstTape, nBodyLegTapePos - 1) ;
   if ( nFoldHt > nExtTempltTol3min && !FootLess) 
	nLegStartHt = FNLegStartHt ( nFirstTape, nBodyLegTapePos - 2)  ;

// Get Origin

if (!GetUser ("xy","Give origin point for Leg", &xyOtemplate)) Exit(%ok, "User Canceled");
Display ("cursor", "wait", "Drawing Second Leg");
// Offset body origin
xyO.y = xyOtemplate.y;
 xyO.x = xyOtemplate.x + nLegStartHt;
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
  }
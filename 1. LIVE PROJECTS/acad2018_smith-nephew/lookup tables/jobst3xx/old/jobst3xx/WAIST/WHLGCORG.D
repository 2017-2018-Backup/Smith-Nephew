// File Name:	WHLGCORG.D
// Purpose:	Establish the origin point to draw the leg
//		and leg template. (CHAP STYLE)
//		This is used when drawing chap style legs and there is no
//		body to be found
//	
// Version:	1.01
// Date:		30.Sep.96
// Author:		Gary George
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
//---------------------------------------------------------------------------------------------
//

if (!GetUser ("xy","Indicate Start Point", &xyOtemplate)) Exit(%ok, "User Canceled");
Display ("cursor", "wait", "Drawing CHAP leg");

PRSetLayerConstruct();
hEnt= AddEntity("marker","xmarker",xyOtemplate,0.125);
if (hEnt) {
  sID=StringMiddle(MakeString("long",UID("get",hEnt)), 1, 4) ; 
  while (StringLength(sID) < 4) sID = sID + " ";
  sID = sID + sFileNo + sLeg ;
  SetDBData (hEnt,"ID",sID + "Origin");
  SetDBData (hEnt,"units", sUnits) ;
  SetDBData (hEnt,"age", sAge) ;
  SetDBData (hEnt,"Data", "CHAP") ;
  SetDBData (hEnt,"Pressure", sPressure) ;
  SetDBData (hEnt,"TapeLengthsPt1",StringMiddle(sTapeLengths, 1, 60));
  SetDBData (hEnt,"TapeLengthsPt2",StringMiddle(sTapeLengths, 61, StringLength(sTapeLengths)-60));
  SetDBData (hEnt,"WaistCir", sWaistCir) ;
  SetDBData (hEnt,"WaistHt", sWaistHt) ;
  SetDBData (hEnt,"TOSCir", sTOSCir) ;
  SetDBData (hEnt,"TOSHt", sTOSHt) ;
  SetDBData (hEnt,"patient", sPatient) ;
  SetDBData (hEnt,"WorkOrder", sWorkOrder) ;
  SetDBData (hEnt,"TOSGivenRed", sTOSRed) ;
  SetDBData (hEnt,"WaistGivenRed", sWaistRed) ;
  SetDBData (hEnt,"FoldHt", sFoldHt) ;
  SetDBData (hEnt,"Body", sLegStyle) ;
  SetDBData (hEnt,"FirstTape", MakeString("scalar", nFirstTape));
  SetDBData (hEnt,"LastTape", MakeString("scalar", nLastTape));
  }

	

	
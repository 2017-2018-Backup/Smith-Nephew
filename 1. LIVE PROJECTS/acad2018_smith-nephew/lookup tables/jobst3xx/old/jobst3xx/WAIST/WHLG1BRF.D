// File Name:	WHLG1BRF.D
// Purpose:	Draw Brief Only
//		ie Back Body Only
//
// Version:	2.01
// Date:		25.Jan.94
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action
// 19.Dec.95	GG	ImageABLE to Triton link


// Subroutines and procedures

procedure PRSetLayerTemplateLeft () {
   hLayer = Table("find", "layer", "TemplateLeft") ;
   if ( hLayer != %badtable )
 	Execute ("menu", "SetLayer",  hLayer) ;
   else
   	Table ("add", "layer",  "TemplateLeft", "Paper Template Left Hand side", "current", Table("find", "color", "Blue"));
}


procedure PRSetLayerConstruct () {
   hLayer = Table("find", "layer", "Construct") ;
   if ( hLayer != %badtable )
 	Execute ("menu", "SetLayer",  hLayer) ;
   else
   	Table ("add", "layer", "Construct", "Construction layer", "current", Table("find", "color", "Gray"));
}

procedure PRSetLayerNotes () {
   hLayer = Table("find", "layer", "Notes") ;
   if ( hLayer != %badtable )
 	Execute ("menu", "SetLayer",  hLayer) ;
   else
   	Table ("add", "layer", "Notes", "Notes in Green", "current", Table("find", "color", "Green"));
}


// Get points
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
		if (StringCompare(sTmp,  sThing+"CO_WaistBott"))
			GetGeometry(hEnt, &sTmp, &xyCO_WaistBott, &xyTmp, &aAngle) ;
			}
		}
	}
   Close("selection", hChan);
 
// Exit if no body origin found
   if (!BodyFound)
	Exit(%cancel, "No BODY found!, draw Cut-Out and try again.\n");   
   
   xyOtemplate.y = xyOBody.y;
   xyOtemplate.x =  xyOBody.x;
   
   // Add a marker at the Origin for later use in drawing the other side
   hEnt = AddEntity("marker","xmarker",xyOtemplate , 0.1, 0.1) ;
   if (hEnt) {
  	sID=StringMiddle(MakeString("long",UID("get",hEnt)), 1, 4) ; 
  	while (StringLength(sID) < 4) sID = sID + " ";
 	sID = sID + sFileNo + sLeg ;
	 }
   SetDBData (hEnt,"ID",sID + "Origin" );
   SetDBData (hEnt,"units", sUnits) ;
// SetDBData (hEnt,"Data", "WH_") ;



// Establish layer
// Always Left in this case
   PRSetLayerTemplateLeft () ;	
   hTemplateLayer = Table("find", "layer", "TemplateLeft");
	
// draw on layer construct
// 'cause drafix is £"^%%%&*&(*&@@ 
   PRSetLayerConstruct();
   // Draw as a panty leg
   StartPoly ("fitted") ;
	AddVertex (xyThigh) ;
	AddVertex (xyFoldPanty) ;
	AddVertex (xyLargest);
	AddVertex (xyMidPoint);
	AddVertex (xyWaist);
	if (Value("scalar",sTOSCir) != 0) AddVertex (xyTOS);	
   EndPoly () ;

   if (Value("scalar",sTOSCir) == 0){				
   	Execute ("menu", "SetLayer", hTemplateLayer);
	AddEntity ("line",  xyWaist, xyTOS) ;
	}

    // Get polyline entity handle
    // Change layer and set DB values

    hChan = Open ("selection", "layer = 'Construct' AND type = 'Curve'");
    if( hChan ) {
	ResetSelection (hChan);
	hCurv = GetNextSelection(hChan);
	SetEntityData ( hCurv, "layer", hTemplateLayer);
   	SetDBData(hCurv, "ID", sID+"LegCurve") ;
	}
   Close ("selection", hChan);

  PRSetLayerNotes() ;
  SetData("TextVertJust", 32);		// Bottom
  SetData("TextHorzJust", 1);		// Left
  SetData("TextHeight", 0.1);
  nTxtY = xyOtemplate.y - 0.5 ;
  nTxtX = xyOtemplate.x + 2;	
	
  if ( StringCompare(sLeg , "Left") ) sText = "Left Leg\n  " ;
	else  sText = "Right Leg\n  ";
 // Work Order Info
  sText = sText + sPatient  + "\n  " + sWorkOrder ;
  if (nFabricClass == 0)  sText = sText + "\n  " + StringMiddle(sFabric,5,StringLength(sFabric) - 4);
	else sText = sText + "\n  " +  sFabric;
   AddEntity("text", sText, nTxtX, nTxtY);	
// Patient info		
   PRSetLayerConstruct() ;
   sText = "  " + sFileNo + "\n  "  + sDiagnosis + "\n  " + sAge + "\n  "  + sSEX  ;
   AddEntity("text", sText, nTxtX, nTxtY-0.75);	

 

// Label at waist
   if (Value("scalar",sTOSCir) == 0){	
	PRSetLayerNotes() ;
   	SetData("TextHorzJust", 4);		// Left
   	SetData("TextHeight", 0.125);
	nTxtY = xyCO_WaistBott.y - 0.5 ;
	nTxtX = xyCO_WaistBott.x - 1;	
	AddEntity ("line", xyCO_WaistBott, nTxtX, nTxtY);
   	hEnt = AddEntity("marker","closed arrow", xyCO_WaistBott , 0.5 ,0.125, Calc("angle",xyCO_WaistBott, nTxtX, nTxtY) ) ;
 			
	AddEntity("text", "TAPER FRONT", nTxtX, nTxtY );
	nTxtY = nTxtY - .15;   
	AddEntity("text", "     TO HERE    ", nTxtX, nTxtY );
	}

   
 Execute ("menu", "SetLayer",  Table("find", "layer", "1")) ;

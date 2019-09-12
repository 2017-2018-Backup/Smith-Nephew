// File Name:	WHLG2BRF.D
// Purpose:	Draw Brief Only
//		ie Second side of brief
//
// Version:	1.01
// Date:		25.Jan.94
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date		By	Action


// Subroutines and procedures

procedure PRSetLayerTemplateLeft () {
   hLayer = Table("find", "layer", "TemplateLeft") ;
   if ( hLayer > %zero )
 	Execute ("menu", "SetLayer",  hLayer) ;
   else
   	Table ("add", "layer", sLayer, "Paper Template Left Hand side", "current", Table("find", "color", "Blue"));
}

procedure PRSetLayerConstruct () {
   hLayer = Table("find", "layer", "Construct") ;
   if ( hLayer > %zero )
 	Execute ("menu", "SetLayer",  hLayer) ;
   else
   	Table ("add", "layer", sLayer, "Construction layer", "current", Table("find", "color", "Gray"));
}

procedure PRSetLayerNotes () {
   hLayer = Table("find", "layer", "Notes") ;
   if ( hLayer > %zero )
 	Execute ("menu", "SetLayer",  hLayer) ;
   else
   	Table ("add", "layer", sLayer, "Notes in Green", "current", Table("find", "color", "Green"));
}

// Get points
   sThing =sFileNo+sFirstLeg ;
   BodyFound = %false ;
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
		if (StringCompare(sTmp,  sThing+"Fold"))
			GetGeometry(hEnt, &sTmp, &xyFold , &xyTmp, &aAngle) ;
		if (StringCompare(sTmp,  sThing+"FoldPanty"))
			GetGeometry(hEnt, &sTmp, &xyFoldPanty , &xyTmp, &aAngle) ;
		if (StringCompare(sTmp,  sThing+"Waist"))
			GetGeometry(hEnt, &sTmp, &xyWaist , &xyTmp, &aAngle) ;
		if (StringCompare(sTmp,  sThing+"CO_WaistBott"))
			GetGeometry(hEnt, &sTmp, &xyCO_WaistBott , &xyTmp, &aAngle) ;
		if (StringCompare(sTmp,  sThing+"CO_ArcCen"))
			GetGeometry(hEnt, &sTmp, &xyCO_ArcCen , &xyTmp, &aAngle) ;
		if (StringCompare(sTmp,  sThing+"Thigh"))
			GetGeometry(hEnt, &sTmp, &xyThigh, &xyTmp, &aAngle) ;
			}
		}
	}
   Close("selection", hChan);
 
// Exit if no body origin found
   if (!BodyFound)
	Exit(%cancel, "No BODY found!, draw Cut-Out and try again.\n");   
   
   xyOtemplate.y = xyOBody.y;
   xyOtemplate.x =  xyOBody.x;
	
// Establish layer
// Always Left in this case
   PRSetLayerTemplateLeft () ;	
   hTemplateLayer = Table("find", "layer", "TemplateLeft");
	
// draw on layer construct
// 'cause drafix is £"^%%%&*&(*&@@ 
   PRSetLayerConstruct();

// Loop through first leg data file, mirror the points found above xyThigh.x
//
   StartPoly ("fitted") ;

   	hFileCurve = Open ("file", "C:\\JOBST\\LEGCURVE.DAT", "readonly") ;
   	ii = 1 ;
   	while ( ii <= nLegVertexCount ) {
		GetLine(hFileCurve, &sLine) ;
		ScanLine(sLine,"blank", &nX, &nY ) ;
		if (nX >= xyThigh.x) {
			xyTmp.x = nX;
			xyTmp.y = nY + 2 * ( xyOBody.y - nY) ;
			AddVertex (  xyTmp ) ;
			}
		if ( ii == 1 )  xyEnd = xyTmp ;
		ii = ii + 1;
		}
   	Close ("file", hFileCurve) ;  

   EndPoly();
  
// Get polyline entity handle
// Change layer and set DB values

    hChan = Open ("selection", "layer = 'Construct' AND type = 'Curve'");
    if( hChan ) {
	ResetSelection (hChan);
	hCurv = GetNextSelection(hChan);
	SetEntityData ( hCurv, "layer", hTemplateLayer);
   	SetDBData(hCurv, "ID", sFileNo+sLeg+"2ndLegCurve") ;
	}
   Close ("selection", hChan);

// Closing lines
   Execute ("menu", "SetLayer", hTemplateLayer);
   if (Value("scalar",sTOSCir) == 0)		
	AddEntity ("line",  	xyWaist.x, xyWaist.y + 2* ( xyOBody.y - xyWaist.y )  ,
			xyTOS.x, xyTOS.y + 2 * ( xyOBody.y - xyTOS.y )  ) ;

   AddEntity ("line", 	xyTOS.x, xyOBody.y ,
		xyTOS.x, xyTOS.y + 2 * ( xyOBody.y - xyTOS.y )  ) ;

   xyPt1.y =  xyCO_ArcCen.y ;
   xyPt1.x =  xyThigh.x ;
   AddEntity ("line", 	xyThigh, xyPt1) ;

   xyPt2.y = xyOBody.y ;
   xyPt2.x = xyFold.x + nFoldOff ;

   AddEntity ("line", xyEnd, xyEnd.x, xyEnd.y + 1.5) ;
   xyEnd.y = xyEnd.y + 1.5 ; 
   
   StartPoly( "fitted");
	AddVertex ( xyPt1) ;
	
	// Sweep curve past cutout
	aAngle = 180 ;
	aAngleInc = 15 ;
	nRadius = xyCO_ArcCen.x - xyThigh.x ;
	ii = 1 ;
	while ( ii  <= 2) {
		AddVertex (CalcXY ("relpolar", xyCO_ArcCen , nRadius, aAngle + aAngleInc * ii));
		ii = ii + 1 ;
		}
	
	AddVertex ( xyPt2) ;

	// Fair into straight line			
	ii = 0 ;
	aAngle = 150;
	nRadius = nRadius / 2 ;
	xyPt1 = CalcXY ("relpolar", xyEnd , nRadius, 0);
	while ( ii  <  2) {
		AddVertex (CalcXY ("relpolar", xyPt1 , nRadius, aAngle + aAngleInc * ii));
		ii = ii + 1 ;
		}

	AddVertex ( xyEnd);

   EndPoly() ;
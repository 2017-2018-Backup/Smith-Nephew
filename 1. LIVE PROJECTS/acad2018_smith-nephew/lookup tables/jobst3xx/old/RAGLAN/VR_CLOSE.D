// File Name:	VR_CLOSE.D
// Purpose:	Draw Bottom of Raglan Curve
//
// Version:	1.01
// Date:		1.Nov.94
// Author:		Gary George
//
//---------------------------------------------------------------------------------------------
// REVISIONS:	
// Date	By	Action
//--------------------------------------------------------------------------------------------------------------------------------------
// -------------------------------------------------- Bottom Raglan  --------------------------------------------------------
//--------------------------------------------------------------------------------------------------------------------------------------
// Draw fold line at raglan curve point
  PRSetLayerNotes() ;
  hEnt = AddEntity( "line",  xyEndBottomCurve.x, xyEndBottomCurve.y + nShouldertoBackRagRedFac, xyRaglanNeck) ;
  SetDBData(hEnt, "ID",sID);
  SetDBData(hEnt, "Data",sData);
  SetDBData(hEnt, "curvetype","sleevecrease");

// Before drawing the curve we must establish the Tangent point to the bottom edge of
// the template.
// We can then use this to calculate a new Origin and Angle for the drawing of the curve
//
   nLowerCurveRadius = Calc ("length", xyLowerAxilla,  xyEndBottomCurve) ;
   xyTangent.x = 0;
   xyTangent.y = 100000 ;
   hCurve = Open ("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly");
   GetLine(hCurve, &sLine);
   ScanLine( sLine, "blank", &nLength, &aAngle);
   aPrevAngle = (Calc("angle",xyLowerAxilla,  xyEndBottomCurve ) - aLowerCurve);      
   xyPt1 = xyLowerAxilla;
   while (GetLine(hCurve,&sLine)) {
	aAngle = aAngle + aPrevAngle;
	xyPt2 = CalcXY("relpolar", xyPt1,nLength, aAngle);
	if ( xyPt2.y >  xyTangent.y) break ;			// break when lowest point passed
	if ( xyPt2.y <=  xyTangent.y) 	xyTangent = xyPt2 ;	// Set initial tangent point
	xyPt1 = xyPt2 ;
	aPrevAngle = aAngle;
	ScanLine( sLine, "blank", &nLength, &aAngle);
	}
   Close ("file", hCurve);

// Having got the initial tangent ( ie the lowest point on the curve) we must use this to
// establish the actual tangent point on the lower edge of the template and get the
// angle to rotate the startpoint of the bottonm curve to

   nError = FN_CirLinInt ( xyLowerAxilla,  
		        xyEndBottomCurve.x, xyLowerAxilla.y, 
		        xyEndBottomCurve, 
		        Calc ( "length", xyTangent,  xyEndBottomCurve)) ;
   aAngle = Calc ("angle", xyEndBottomCurve, xyTangent) - Calc ("angle", xyEndBottomCurve, xyInt) ;
   xyStartBottomCurve = CalcXY ("relpolar",  
			      xyEndBottomCurve,
			      Calc ("length", xyEndBottomCurve, xyLowerAxilla),
			      Calc ("angle", xyEndBottomCurve, xyLowerAxilla) - aAngle) ;
   xyTangent = xyInt ;

//
// DRAW BOTTOM CURVE
//
// set layer

    if (StringCompare(sSleeve, "Right"))
	PRSetLayerTemplateRight() ;
    else
	PRSetLayerTemplateLeft() ;
//
   hCurve = Open ("file", sPathJOBST + "\\TEMPLTS\\VESTCURV.DAT", "readonly");
   GetLine(hCurve, &sLine);
   ScanLine( sLine, "blank", &nLength, &aAngle);
   // Rotate curve to correct start angle
   aPrevAngle = (Calc("angle", xyStartBottomCurve,  xyEndBottomCurve ) - aLowerCurve);     
   xyPt1 = xyStartBottomCurve;
   DrawSegment = %false ;		// Set draw segments flag off
   while (GetLine(hCurve, &sLine)){
	aAngle = aAngle + aPrevAngle;
	xyPt2 = CalcXY("relpolar", xyPt1,nLength, aAngle);
	if ( FN_CirLinInt( xyPt1, xyPt2, xyLowerAxilla, nLowerCurveRadius)) {
		AddVertex(xyEndBottomCurve);
		EndPoly();
          	hEnt = UID ("find", UID ("getmax")) ;
    		SetDBData(hEnt, "ID",sID);
    		SetDBData(hEnt, "curvetype","sleeveraglanbottom");
    		SetDBData(hEnt, "Data", sData);
		break;
		}
	// Pick up curve after it has passed tangent point
	if (xyPt2.x > xyTangent.x && DrawSegment == %false) {
		DrawSegment = %true ;
		xyPt1 = xyTangent ;
		StartPoly("polyline");
		AddVertex(xyLowerAxilla);
		AddVertex(xyPt1);
		}

	if ( DrawSegment ) AddVertex(xyPt2);
	xyPt1 = xyPt2 ;
	aPrevAngle = aAngle;
	ScanLine( sLine, "blank", &nLength, &aAngle);
   } //End GetLine while
  Close ("file", hCurve);

// Draw Closing lines
  hEnt = AddEntity ("line", xyEndBottomCurve, xyBackNeck);
  SetDBData(hEnt, "ID",sID);
  SetDBData(hEnt, "curvetype","sleeveclosing");
  SetDBData(hEnt, "Data", sData);

// Reset to layer 1
  PRSetLayer1 () ;